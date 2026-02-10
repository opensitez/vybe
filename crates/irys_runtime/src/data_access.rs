//! Data Access layer for irys.
//!
//! Provides a unified backend that supports both ADODB (VB6) and ADO.NET
//! (VB.NET) API surfaces, backed by sqlx for SQLite, PostgreSQL, and MySQL.
//!
//! Architecture:
//!   VB Code (ADODB or ADO.NET syntax)
//!       ↓
//!   Interpreter dispatch
//!       ↓
//!   data_access.rs  (this module)
//!       ↓
//!   sqlx + block_on(tokio)
//!       ↓
//!   SQLite / PostgreSQL / MySQL

use std::collections::HashMap;
use std::sync::{Arc, Mutex, atomic::{AtomicU64, Ordering}};

/// Global connection ID counter.
static NEXT_CONN_ID: AtomicU64 = AtomicU64::new(1);
/// Global recordset/reader ID counter.
static NEXT_RS_ID: AtomicU64 = AtomicU64::new(1);

/// Detected database backend from connection string.
#[derive(Debug, Clone, PartialEq)]
pub enum DbBackend {
    Sqlite,
    Postgres,
    MySql,
}

/// A row of results: column names → values (all stored as strings for simplicity).
#[derive(Debug, Clone)]
pub struct DbRow {
    pub columns: Vec<String>,
    pub values: Vec<String>,
}

impl DbRow {
    /// Get a value by column name (case-insensitive).
    pub fn get_by_name(&self, name: &str) -> Option<&str> {
        let lower = name.to_lowercase();
        for (i, col) in self.columns.iter().enumerate() {
            if col.to_lowercase() == lower {
                return self.values.get(i).map(|s| s.as_str());
            }
        }
        None
    }

    /// Get a value by ordinal index.
    pub fn get_by_index(&self, index: usize) -> Option<&str> {
        self.values.get(index).map(|s| s.as_str())
    }
}

/// Holds the results of a query — a vector of rows plus current position.
#[derive(Debug, Clone)]
pub struct RecordSet {
    pub id: u64,
    pub columns: Vec<String>,
    pub rows: Vec<DbRow>,
    pub position: i32,          // -1 = BOF, 0..n = current row, n = EOF
    pub records_affected: i64,  // For INSERT/UPDATE/DELETE
}

impl RecordSet {
    pub fn new(columns: Vec<String>, rows: Vec<DbRow>, records_affected: i64) -> Self {
        let pos = if rows.is_empty() { 0 } else { 0 };
        Self {
            id: NEXT_RS_ID.fetch_add(1, Ordering::SeqCst),
            columns,
            rows,
            position: pos,
            records_affected,
        }
    }

    pub fn eof(&self) -> bool {
        self.rows.is_empty() || self.position >= self.rows.len() as i32
    }

    pub fn bof(&self) -> bool {
        self.position < 0
    }

    pub fn move_next(&mut self) {
        self.position += 1;
    }

    pub fn move_first(&mut self) {
        if !self.rows.is_empty() {
            self.position = 0;
        }
    }

    pub fn move_last(&mut self) {
        if !self.rows.is_empty() {
            self.position = self.rows.len() as i32 - 1;
        }
    }

    pub fn move_previous(&mut self) {
        self.position -= 1;
    }

    /// Get current row (None if EOF/BOF).
    pub fn current_row(&self) -> Option<&DbRow> {
        if self.position >= 0 && (self.position as usize) < self.rows.len() {
            Some(&self.rows[self.position as usize])
        } else {
            None
        }
    }

    /// ADO.NET-style Read(): advances to next row, returns true if a row is available.
    /// On first call, stays at row 0; subsequent calls advance.
    pub fn read(&mut self) -> bool {
        if self.position < 0 {
            self.position = 0;
        } else {
            self.position += 1;
        }
        !self.eof()
    }

    pub fn record_count(&self) -> i32 {
        self.rows.len() as i32
    }

    pub fn field_count(&self) -> i32 {
        self.columns.len() as i32
    }
}

/// Manages database connections and provides query execution.
///
/// Uses a tokio Runtime internally to drive sqlx's async API synchronously.
/// The runtime lives on a dedicated background thread so it can be used even
/// when the caller is already inside a tokio runtime (e.g. Dioxus desktop).
pub struct DataAccessManager {
    /// Sender to dispatch async work to the background runtime thread.
    bg_sender: std::sync::mpsc::Sender<Box<dyn FnOnce(&tokio::runtime::Runtime) + Send>>,
    /// Active connections keyed by connection ID.
    connections: HashMap<u64, DbConnection>,
    /// Active recordsets keyed by recordset ID.
    pub recordsets: HashMap<u64, RecordSet>,
    /// Multi-result storage: rs_id → vec of additional RecordSets.
    pub pending_results: HashMap<u64, Vec<RecordSet>>,
}

struct DbConnection {
    #[allow(dead_code)]
    id: u64,
    backend: DbBackend,
    pool: sqlx::AnyPool,
    #[allow(dead_code)]
    connection_string: String,
    /// Connection timeout parsed from connection string (seconds).
    #[allow(dead_code)]
    pub connect_timeout: u32,
    /// Max pool size parsed from connection string.
    #[allow(dead_code)]
    pub max_pool_size: u32,
}

impl DataAccessManager {
    pub fn new() -> Self {
        // Spawn a dedicated thread that owns the tokio runtime.
        // This avoids "Cannot start a runtime from within a runtime" when
        // the caller (e.g. Dioxus/editor) already has an active tokio context.
        let (tx, rx) = std::sync::mpsc::channel::<Box<dyn FnOnce(&tokio::runtime::Runtime) + Send>>();
        std::thread::Builder::new()
            .name("irys-data-rt".into())
            .spawn(move || {
                let rt = tokio::runtime::Builder::new_current_thread()
                    .enable_all()
                    .build()
                    .expect("Failed to create tokio runtime for data access");
                while let Ok(task) = rx.recv() {
                    task(&rt);
                }
            })
            .expect("Failed to spawn data-access runtime thread");

        Self {
            bg_sender: tx,
            connections: HashMap::new(),
            recordsets: HashMap::new(),
            pending_results: HashMap::new(),
        }
    }

    /// Run an async block on the background tokio runtime and wait for the
    /// result.  Safe to call from inside another tokio runtime.
    fn block_on_bg<F, R>(&self, f: F) -> R
    where
        F: FnOnce(&tokio::runtime::Runtime) -> R + Send + 'static,
        R: Send + 'static,
    {
        let (done_tx, done_rx) = std::sync::mpsc::channel::<R>();
        self.bg_sender
            .send(Box::new(move |rt| {
                let result = f(rt);
                let _ = done_tx.send(result);
            }))
            .expect("Data-access runtime thread has exited");
        done_rx.recv().expect("Data-access runtime thread panicked")
    }

    /// Parse a connection string and determine the backend.
    ///
    /// Supported formats:
    ///   SQLite:    "Data Source=mydb.sqlite" or "sqlite:mydb.sqlite" or "Data Source=:memory:"
    ///   Postgres:  "Host=localhost;Database=mydb;Username=user;Password=pass" or "postgres://user:pass@host/db"
    ///   MySQL:     "Server=localhost;Database=mydb;Uid=user;Pwd=pass" or "mysql://user:pass@host/db"
    ///   ADODB:     "Provider=Microsoft.Jet.OLEDB..." (mapped to SQLite)
    ///              "Provider=SQLOLEDB;..." (mapped to MySQL/MSSQL-compatible)
    fn parse_connection_string(conn_str: &str) -> Result<(DbBackend, String), String> {
        let trimmed = conn_str.trim();
        let lower = trimmed.to_lowercase();

        // Direct URL format: sqlite:, postgres://, mysql://
        if lower.starts_with("sqlite:") {
            let path = &trimmed[7..];
            return Ok((DbBackend::Sqlite, format!("sqlite:{}", path)));
        }
        if lower.starts_with("postgres://") || lower.starts_with("postgresql://") {
            return Ok((DbBackend::Postgres, trimmed.to_string()));
        }
        if lower.starts_with("mysql://") {
            return Ok((DbBackend::MySql, trimmed.to_string()));
        }

        // Parse key=value pairs (ADO/ADO.NET style)
        let pairs = Self::parse_kv_pairs(trimmed);

        // Check Provider for ADODB
        let provider = pairs.get("provider").map(|s| s.to_lowercase()).unwrap_or_default();

        // SQLite detection
        if provider.contains("jet") || provider.contains("ace")
            || lower.contains("data source=:memory:")
            || pairs.get("data source").map(|s| {
                let sl = s.to_lowercase();
                sl.ends_with(".sqlite") || sl.ends_with(".db") || sl.ends_with(".sqlite3")
                    || sl.ends_with(".mdb") || sl.ends_with(".accdb") || sl == ":memory:"
            }).unwrap_or(false)
        {
            let ds = pairs.get("data source").cloned().unwrap_or(":memory:".to_string());
            return Ok((DbBackend::Sqlite, format!("sqlite:{}", ds)));
        }

        // PostgreSQL detection
        if provider.contains("postgre") || pairs.contains_key("host") && !pairs.contains_key("server") {
            let host = pairs.get("host").cloned().unwrap_or("localhost".to_string());
            let port = pairs.get("port").cloned().unwrap_or("5432".to_string());
            let db = pairs.get("database").cloned().unwrap_or("postgres".to_string());
            let user = pairs.get("username").or(pairs.get("user id")).or(pairs.get("uid")).cloned().unwrap_or("postgres".to_string());
            let pass = pairs.get("password").or(pairs.get("pwd")).cloned().unwrap_or_default();
            return Ok((DbBackend::Postgres, format!("postgres://{}:{}@{}:{}/{}", user, pass, host, port, db)));
        }

        // MySQL/MSSQL detection (Server= syntax)
        if provider.contains("sqloledb") || provider.contains("mysql")
            || pairs.contains_key("server")
        {
            let server = pairs.get("server").cloned().unwrap_or("localhost".to_string());
            let port = pairs.get("port").cloned().unwrap_or("3306".to_string());
            let db = pairs.get("database").or(pairs.get("initial catalog")).cloned().unwrap_or("mysql".to_string());
            let user = pairs.get("uid").or(pairs.get("user id")).or(pairs.get("username")).cloned().unwrap_or("root".to_string());
            let pass = pairs.get("pwd").or(pairs.get("password")).cloned().unwrap_or_default();
            return Ok((DbBackend::MySql, format!("mysql://{}:{}@{}:{}/{}", user, pass, server, port, db)));
        }

        // Fallback: treat as SQLite with the whole string as a path
        if !trimmed.is_empty() {
            return Ok((DbBackend::Sqlite, format!("sqlite:{}", trimmed)));
        }

        Err("Cannot determine database backend from connection string".to_string())
    }

    /// Parse "Key=Value;Key2=Value2" into a HashMap.
    fn parse_kv_pairs(s: &str) -> HashMap<String, String> {
        let mut map = HashMap::new();
        for part in s.split(';') {
            let part = part.trim();
            if let Some(eq_pos) = part.find('=') {
                let key = part[..eq_pos].trim().to_lowercase();
                let val = part[eq_pos+1..].trim().to_string();
                map.insert(key, val);
            }
        }
        map
    }

    /// Open a connection. Returns a connection ID.
    pub fn open_connection(&mut self, conn_str: &str) -> Result<u64, String> {
        let (backend, url) = Self::parse_connection_string(conn_str)?;

        // Parse pool and timeout settings from connection string
        let pairs = Self::parse_kv_pairs(conn_str);
        let max_pool: u32 = pairs.get("max pool size")
            .or(pairs.get("maxpoolsize"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(5);
        let min_pool: u32 = pairs.get("min pool size")
            .or(pairs.get("minpoolsize"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(0);
        let conn_timeout: u32 = pairs.get("connect timeout")
            .or(pairs.get("connection timeout"))
            .or(pairs.get("timeout"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(30);

        // Install the sqlx Any driver for this backend
        sqlx::any::install_default_drivers();

        let backend_clone = backend.clone();
        let url_clone = url.clone();
        let pool = self.block_on_bg(move |rt| {
            rt.block_on(async {
                let mut opts = sqlx::any::AnyPoolOptions::new();
                if backend_clone == DbBackend::Sqlite && url_clone.contains(":memory:") {
                    opts = opts.max_connections(1).min_connections(1);
                } else {
                    opts = opts.max_connections(max_pool).min_connections(min_pool);
                }
                opts = opts.acquire_timeout(std::time::Duration::from_secs(conn_timeout as u64));
                opts.connect(&url_clone)
                    .await
            })
        }).map_err(|e| format!("Failed to connect: {}", e))?;

        let id = NEXT_CONN_ID.fetch_add(1, Ordering::SeqCst);
        self.connections.insert(id, DbConnection {
            id,
            backend,
            pool,
            connection_string: conn_str.to_string(),
            connect_timeout: conn_timeout,
            max_pool_size: max_pool,
        });

        Ok(id)
    }

    /// Close a connection.
    pub fn close_connection(&mut self, conn_id: u64) -> Result<(), String> {
        if let Some(conn) = self.connections.remove(&conn_id) {
            let pool = conn.pool.clone();
            self.block_on_bg(move |rt| {
                rt.block_on(async move {
                    pool.close().await;
                });
            });
            Ok(())
        } else {
            Err(format!("Connection {} not found", conn_id))
        }
    }

    /// Execute a query that returns rows (SELECT).
    /// Returns a RecordSet ID.
    pub fn execute_reader(&mut self, conn_id: u64, sql: &str) -> Result<u64, String> {
        let conn = self.connections.get(&conn_id)
            .ok_or_else(|| format!("Connection {} not found or not open", conn_id))?;

        let pool = conn.pool.clone();
        let sql_owned = sql.to_string();
        let rows_result = self.block_on_bg(move |rt| {
            rt.block_on(async move {
                use sqlx::Row;
                use sqlx::Column;

                let raw_rows = sqlx::query(&sql_owned)
                    .fetch_all(&pool)
                    .await
                    .map_err(|e| format!("Query error: {}", e))?;

                let mut columns: Vec<String> = Vec::new();
                let mut db_rows: Vec<DbRow> = Vec::new();

                for raw_row in &raw_rows {
                    if columns.is_empty() {
                        columns = raw_row.columns().iter()
                            .map(|c| c.name().to_string())
                            .collect();
                    }

                    let mut values = Vec::new();
                    for i in 0..raw_row.columns().len() {
                        let val: String = raw_row.try_get::<String, _>(i)
                            .or_else(|_| raw_row.try_get::<i64, _>(i).map(|v| v.to_string()))
                            .or_else(|_| raw_row.try_get::<f64, _>(i).map(|v| v.to_string()))
                            .or_else(|_| raw_row.try_get::<bool, _>(i).map(|v| v.to_string()))
                            .unwrap_or_else(|_| "NULL".to_string());
                        values.push(val);
                    }

                    db_rows.push(DbRow {
                        columns: columns.clone(),
                        values,
                    });
                }

                Ok::<(Vec<String>, Vec<DbRow>), String>((columns, db_rows))
            })
        })?;

        let (columns, db_rows) = rows_result;
        let rs = RecordSet::new(columns, db_rows, 0);
        let rs_id = rs.id;
        self.recordsets.insert(rs_id, rs);

        Ok(rs_id)
    }

    /// Execute a non-query statement (INSERT, UPDATE, DELETE, CREATE TABLE, etc.).
    /// Returns the number of rows affected.
    pub fn execute_non_query(&mut self, conn_id: u64, sql: &str) -> Result<i64, String> {
        let conn = self.connections.get(&conn_id)
            .ok_or_else(|| format!("Connection {} not found or not open", conn_id))?;

        let pool = conn.pool.clone();
        let sql_owned = sql.to_string();
        let result = self.block_on_bg(move |rt| {
            rt.block_on(async move {
                sqlx::query(&sql_owned)
                    .execute(&pool)
                    .await
                    .map_err(|e| format!("Execute error: {}", e))
            })
        })?;

        Ok(result.rows_affected() as i64)
    }

    /// Execute SQL — auto-detects SELECT vs non-SELECT.
    /// For SELECT: returns a RecordSet ID (positive).
    /// For non-SELECT: returns rows affected as negative (to distinguish).
    pub fn execute(&mut self, conn_id: u64, sql: &str) -> Result<ExecuteResult, String> {
        let trimmed = sql.trim();
        let upper = trimmed.to_uppercase();
        if upper.starts_with("SELECT") || upper.starts_with("PRAGMA") || upper.starts_with("SHOW")
            || upper.starts_with("DESCRIBE") || upper.starts_with("EXPLAIN")
        {
            let rs_id = self.execute_reader(conn_id, sql)?;
            Ok(ExecuteResult::Recordset(rs_id))
        } else {
            let affected = self.execute_non_query(conn_id, sql)?;
            Ok(ExecuteResult::RowsAffected(affected))
        }
    }

    /// Execute a scalar query — returns the first column of the first row.
    pub fn execute_scalar(&mut self, conn_id: u64, sql: &str) -> Result<String, String> {
        let rs_id = self.execute_reader(conn_id, sql)?;
        let result = if let Some(rs) = self.recordsets.get(&rs_id) {
            if let Some(row) = rs.rows.first() {
                row.values.first().cloned().unwrap_or("NULL".to_string())
            } else {
                "NULL".to_string()
            }
        } else {
            "NULL".to_string()
        };
        self.recordsets.remove(&rs_id);
        Ok(result)
    }

    /// Close/remove a recordset.
    pub fn close_recordset(&mut self, rs_id: u64) {
        self.recordsets.remove(&rs_id);
    }

    /// Check if a connection is open.
    pub fn is_connected(&self, conn_id: u64) -> bool {
        self.connections.contains_key(&conn_id)
    }

    /// Get the backend type for a connection.
    pub fn get_backend(&self, conn_id: u64) -> Option<DbBackend> {
        self.connections.get(&conn_id).map(|c| c.backend.clone())
    }

    // ===== Transactions =====

    /// Begin a transaction.  Returns a transaction ID.
    pub fn begin_transaction(&mut self, conn_id: u64) -> Result<u64, String> {
        let conn = self.connections.get(&conn_id)
            .ok_or_else(|| format!("Connection {} not found", conn_id))?;
        let pool = conn.pool.clone();
        self.block_on_bg(move |rt| {
            rt.block_on(async move {
                sqlx::query("BEGIN").execute(&pool).await
                    .map_err(|e| format!("BEGIN error: {}", e))
            })
        })?;
        // Use conn_id as transaction id (one active tx per connection)
        Ok(conn_id)
    }

    /// Commit a transaction.
    pub fn commit(&mut self, conn_id: u64) -> Result<(), String> {
        let conn = self.connections.get(&conn_id)
            .ok_or_else(|| format!("Connection {} not found", conn_id))?;
        let pool = conn.pool.clone();
        self.block_on_bg(move |rt| {
            rt.block_on(async move {
                sqlx::query("COMMIT").execute(&pool).await
                    .map_err(|e| format!("COMMIT error: {}", e))
            })
        })?;
        Ok(())
    }

    /// Rollback a transaction.
    pub fn rollback(&mut self, conn_id: u64) -> Result<(), String> {
        let conn = self.connections.get(&conn_id)
            .ok_or_else(|| format!("Connection {} not found", conn_id))?;
        let pool = conn.pool.clone();
        self.block_on_bg(move |rt| {
            rt.block_on(async move {
                sqlx::query("ROLLBACK").execute(&pool).await
                    .map_err(|e| format!("ROLLBACK error: {}", e))
            })
        })?;
        Ok(())
    }

    // ===== Parameterised queries =====

    /// Substitute `@param` placeholders in SQL with literal values.
    /// Parameters is a list of (name, value) where name starts with `@`.
    /// Values are escaped for safe embedding.
    pub fn substitute_params(sql: &str, params: &[(String, String)]) -> String {
        let mut result = sql.to_string();
        // Sort by name length descending so @param10 is replaced before @param1
        let mut sorted: Vec<&(String, String)> = params.iter().collect();
        sorted.sort_by(|a, b| b.0.len().cmp(&a.0.len()));
        for (name, value) in sorted {
            let escaped = value.replace('\'', "''");
            // Try to keep numeric values unquoted
            let replacement = if value.parse::<f64>().is_ok() || value.eq_ignore_ascii_case("null") {
                value.clone()
            } else {
                format!("'{}'", escaped)
            };
            result = result.replace(name.as_str(), &replacement);
        }
        result
    }

    // ===== Stored Procedures =====

    /// Execute a stored procedure.
    /// Wraps the procedure name into the appropriate backend-specific CALL syntax.
    pub fn execute_stored_proc(&mut self, conn_id: u64, proc_name: &str, params: &[(String, String)]) -> Result<ExecuteResult, String> {
        let backend = self.connections.get(&conn_id)
            .map(|c| c.backend.clone())
            .ok_or_else(|| format!("Connection {} not found", conn_id))?;

        let call_sql = match backend {
            DbBackend::Sqlite => {
                // SQLite doesn't have stored procedures — just execute as a query
                // This allows users to test with SELECT-based "procedures"
                if params.is_empty() {
                    proc_name.to_string()
                } else {
                    let placeholders: Vec<String> = params.iter().map(|(_, v)| {
                        if v.parse::<f64>().is_ok() || v.eq_ignore_ascii_case("null") {
                            v.clone()
                        } else {
                            format!("'{}'", v.replace('\'', "''"))
                        }
                    }).collect();
                    format!("SELECT {}({})", proc_name, placeholders.join(", "))
                }
            }
            DbBackend::Postgres => {
                if params.is_empty() {
                    format!("SELECT * FROM {}()", proc_name)
                } else {
                    let placeholders: Vec<String> = params.iter().map(|(_, v)| {
                        if v.parse::<f64>().is_ok() || v.eq_ignore_ascii_case("null") {
                            v.clone()
                        } else {
                            format!("'{}'", v.replace('\'', "''"))
                        }
                    }).collect();
                    format!("SELECT * FROM {}({})", proc_name, placeholders.join(", "))
                }
            }
            DbBackend::MySql => {
                if params.is_empty() {
                    format!("CALL {}()", proc_name)
                } else {
                    let placeholders: Vec<String> = params.iter().map(|(_, v)| {
                        if v.parse::<f64>().is_ok() || v.eq_ignore_ascii_case("null") {
                            v.clone()
                        } else {
                            format!("'{}'", v.replace('\'', "''"))
                        }
                    }).collect();
                    format!("CALL {}({})", proc_name, placeholders.join(", "))
                }
            }
        };

        self.execute(conn_id, &call_sql)
    }

    // ===== Multiple Result Sets =====

    /// Execute multiple statements separated by semicolons.
    /// Returns the first result set ID and stores subsequent ones for NextResult().
    pub fn execute_multi(&mut self, conn_id: u64, sql: &str) -> Result<u64, String> {
        // Split SQL by semicolons (respecting quoted strings)
        let statements = Self::split_sql_statements(sql);

        if statements.is_empty() {
            return Err("No SQL statements provided".to_string());
        }

        // Execute first statement as the primary result
        let first_rs_id = self.execute_reader(conn_id, &statements[0])?;

        // Execute remaining and store as pending
        let mut pending = Vec::new();
        for stmt in &statements[1..] {
            let trimmed = stmt.trim();
            if trimmed.is_empty() { continue; }
            let upper = trimmed.to_uppercase();
            if upper.starts_with("SELECT") || upper.starts_with("PRAGMA") || upper.starts_with("SHOW") {
                let rs_id = self.execute_reader(conn_id, trimmed)?;
                if let Some(rs) = self.recordsets.remove(&rs_id) {
                    pending.push(rs);
                }
            }
        }

        if !pending.is_empty() {
            self.pending_results.insert(first_rs_id, pending);
        }

        Ok(first_rs_id)
    }

    /// Advance to next result set. Returns true if another result set is available.
    pub fn next_result(&mut self, rs_id: u64) -> bool {
        if let Some(pending) = self.pending_results.get_mut(&rs_id) {
            if let Some(next_rs) = pending.first().cloned() {
                pending.remove(0);
                // Replace current recordset with next one
                let mut new_rs = next_rs;
                new_rs.id = rs_id; // Reuse the same ID
                self.recordsets.insert(rs_id, new_rs);
                return true;
            }
            self.pending_results.remove(&rs_id);
        }
        false
    }

    /// Split SQL string by semicolons, respecting single-quoted strings.
    fn split_sql_statements(sql: &str) -> Vec<String> {
        let mut stmts = Vec::new();
        let mut current = String::new();
        let mut in_string = false;
        let mut chars = sql.chars().peekable();

        while let Some(ch) = chars.next() {
            if ch == '\'' {
                in_string = !in_string;
                current.push(ch);
            } else if ch == ';' && !in_string {
                let trimmed = current.trim().to_string();
                if !trimmed.is_empty() {
                    stmts.push(trimmed);
                }
                current.clear();
            } else {
                current.push(ch);
            }
        }
        let trimmed = current.trim().to_string();
        if !trimmed.is_empty() {
            stmts.push(trimmed);
        }
        stmts
    }

    // ===== Schema Discovery =====

    /// Get schema information for a connection.
    /// Returns column info as a RecordSet with schema metadata.
    pub fn get_schema(&mut self, conn_id: u64, collection: &str) -> Result<u64, String> {
        let conn = self.connections.get(&conn_id)
            .ok_or_else(|| format!("Connection {} not found", conn_id))?;

        let sql = match conn.backend {
            DbBackend::Sqlite => {
                match collection.to_lowercase().as_str() {
                    "tables" => "SELECT name AS TABLE_NAME, type AS TABLE_TYPE FROM sqlite_master WHERE type IN ('table','view') ORDER BY name".to_string(),
                    "columns" => "PRAGMA table_list".to_string(),
                    _ => format!("SELECT name AS TABLE_NAME, type AS TABLE_TYPE FROM sqlite_master WHERE type='table' ORDER BY name"),
                }
            }
            DbBackend::Postgres => {
                match collection.to_lowercase().as_str() {
                    "tables" => "SELECT table_name AS TABLE_NAME, table_type AS TABLE_TYPE FROM information_schema.tables WHERE table_schema='public' ORDER BY table_name".to_string(),
                    "columns" => "SELECT table_name AS TABLE_NAME, column_name AS COLUMN_NAME, data_type AS DATA_TYPE, is_nullable AS IS_NULLABLE FROM information_schema.columns WHERE table_schema='public' ORDER BY table_name, ordinal_position".to_string(),
                    _ => "SELECT table_name AS TABLE_NAME FROM information_schema.tables WHERE table_schema='public' ORDER BY table_name".to_string(),
                }
            }
            DbBackend::MySql => {
                match collection.to_lowercase().as_str() {
                    "tables" => "SELECT TABLE_NAME, TABLE_TYPE FROM information_schema.tables WHERE TABLE_SCHEMA=DATABASE() ORDER BY TABLE_NAME".to_string(),
                    "columns" => "SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE, IS_NULLABLE FROM information_schema.columns WHERE TABLE_SCHEMA=DATABASE() ORDER BY TABLE_NAME, ORDINAL_POSITION".to_string(),
                    _ => "SELECT TABLE_NAME FROM information_schema.tables WHERE TABLE_SCHEMA=DATABASE() ORDER BY TABLE_NAME".to_string(),
                }
            }
        };

        self.execute_reader(conn_id, &sql)
    }

    /// Get schema table for a reader/recordset — returns column metadata.
    pub fn get_schema_table(&mut self, rs_id: u64) -> Result<u64, String> {
        let columns = if let Some(rs) = self.recordsets.get(&rs_id) {
            rs.columns.clone()
        } else {
            return Err("RecordSet not found".to_string());
        };

        // Build a schema recordset with column info
        let schema_cols = vec![
            "ColumnName".to_string(),
            "ColumnOrdinal".to_string(),
            "DataType".to_string(),
        ];
        let mut schema_rows = Vec::new();
        for (i, col) in columns.iter().enumerate() {
            schema_rows.push(DbRow {
                columns: schema_cols.clone(),
                values: vec![col.clone(), i.to_string(), "String".to_string()],
            });
        }

        let rs = RecordSet::new(schema_cols, schema_rows, 0);
        let id = rs.id;
        self.recordsets.insert(id, rs);
        Ok(id)
    }

    /// Get connection info (provider, server version, etc.)
    pub fn get_connection_info(&self, conn_id: u64) -> Option<(String, String)> {
        self.connections.get(&conn_id).map(|c| {
            let provider = match c.backend {
                DbBackend::Sqlite => "SQLite".to_string(),
                DbBackend::Postgres => "PostgreSQL".to_string(),
                DbBackend::MySql => "MySQL".to_string(),
            };
            (provider, c.connection_string.clone())
        })
    }
}

/// Result of Execute() — either a recordset or rows-affected count.
pub enum ExecuteResult {
    Recordset(u64),
    RowsAffected(i64),
}

/// Thread-safe global data access manager (lazy init).
use std::sync::OnceLock;

static GLOBAL_DAM: OnceLock<Arc<Mutex<DataAccessManager>>> = OnceLock::new();

pub fn get_global_dam() -> Arc<Mutex<DataAccessManager>> {
    GLOBAL_DAM.get_or_init(|| {
        Arc::new(Mutex::new(DataAccessManager::new()))
    }).clone()
}

/// Test a connection string and return a list of table names.
/// This is a convenience function for the editor's properties panel.
/// Returns Ok(vec of table names) on success or Err(error message).
pub fn test_connection_and_list_tables(conn_str: &str) -> Result<Vec<String>, String> {
    let dam_arc = get_global_dam();
    let mut dam = dam_arc.lock().map_err(|e| format!("Lock error: {}", e))?;
    let conn_id = dam.open_connection(conn_str)?;
    let rs_id = dam.get_schema(conn_id, "tables")?;
    let tables: Vec<String> = if let Some(rs) = dam.recordsets.get(&rs_id) {
        rs.rows.iter().filter_map(|row| {
            row.get_by_name("TABLE_NAME").map(|s| s.to_string())
        }).collect()
    } else {
        Vec::new()
    };
    dam.recordsets.remove(&rs_id);
    // Don't close — keep the connection around for reuse by the runtime
    Ok(tables)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_sqlite_connection_string() {
        let (backend, url) = DataAccessManager::parse_connection_string("Data Source=test.db").unwrap();
        assert_eq!(backend, DbBackend::Sqlite);
        assert!(url.starts_with("sqlite:"));
    }

    #[test]
    fn test_parse_sqlite_direct() {
        let (backend, url) = DataAccessManager::parse_connection_string("sqlite:test.db").unwrap();
        assert_eq!(backend, DbBackend::Sqlite);
        assert_eq!(url, "sqlite:test.db");
    }

    #[test]
    fn test_parse_postgres_connection_string() {
        let (backend, url) = DataAccessManager::parse_connection_string(
            "Host=localhost;Database=mydb;Username=user;Password=pass"
        ).unwrap();
        assert_eq!(backend, DbBackend::Postgres);
        assert!(url.starts_with("postgres://"));
    }

    #[test]
    fn test_parse_mysql_connection_string() {
        let (backend, url) = DataAccessManager::parse_connection_string(
            "Server=localhost;Database=mydb;Uid=root;Pwd=secret"
        ).unwrap();
        assert_eq!(backend, DbBackend::MySql);
        assert!(url.starts_with("mysql://"));
    }

    #[test]
    fn test_parse_adodb_jet() {
        let (backend, _url) = DataAccessManager::parse_connection_string(
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=mydb.mdb"
        ).unwrap();
        assert_eq!(backend, DbBackend::Sqlite);
    }
}
