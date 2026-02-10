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
pub struct DataAccessManager {
    runtime: tokio::runtime::Runtime,
    /// Active connections keyed by connection ID.
    connections: HashMap<u64, DbConnection>,
    /// Active recordsets keyed by recordset ID.
    pub recordsets: HashMap<u64, RecordSet>,
}

struct DbConnection {
    #[allow(dead_code)]
    id: u64,
    backend: DbBackend,
    pool: sqlx::AnyPool,
    #[allow(dead_code)]
    connection_string: String,
}

impl DataAccessManager {
    pub fn new() -> Self {
        let runtime = tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .expect("Failed to create tokio runtime for data access");

        Self {
            runtime,
            connections: HashMap::new(),
            recordsets: HashMap::new(),
        }
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

        // Install the sqlx Any driver for this backend
        match backend {
            DbBackend::Sqlite => { sqlx::any::install_default_drivers(); }
            DbBackend::Postgres => { sqlx::any::install_default_drivers(); }
            DbBackend::MySql => { sqlx::any::install_default_drivers(); }
        }

        let pool = self.runtime.block_on(async {
            sqlx::any::AnyPoolOptions::new()
                .max_connections(5)
                .connect(&url)
                .await
        }).map_err(|e| format!("Failed to connect: {}", e))?;

        let id = NEXT_CONN_ID.fetch_add(1, Ordering::SeqCst);
        self.connections.insert(id, DbConnection {
            id,
            backend,
            pool,
            connection_string: conn_str.to_string(),
        });

        Ok(id)
    }

    /// Close a connection.
    pub fn close_connection(&mut self, conn_id: u64) -> Result<(), String> {
        if let Some(conn) = self.connections.remove(&conn_id) {
            self.runtime.block_on(async {
                conn.pool.close().await;
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

        let rows_result = self.runtime.block_on(async {
            use sqlx::Row;
            use sqlx::Column;

            let raw_rows = sqlx::query(sql)
                .fetch_all(&conn.pool)
                .await
                .map_err(|e| format!("Query error: {}", e))?;

            let mut columns: Vec<String> = Vec::new();
            let mut db_rows: Vec<DbRow> = Vec::new();

            for raw_row in &raw_rows {
                // Get column names from the first row
                if columns.is_empty() {
                    columns = raw_row.columns().iter()
                        .map(|c| c.name().to_string())
                        .collect();
                }

                let mut values = Vec::new();
                for i in 0..raw_row.columns().len() {
                    // Try to get as string — sqlx Any backend
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

        let result = self.runtime.block_on(async {
            sqlx::query(sql)
                .execute(&conn.pool)
                .await
                .map_err(|e| format!("Execute error: {}", e))
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
