use std::sync::Arc;
use vybe_runtime::Value;

/// A database driver — one implementation per backend (SQLite, PostgreSQL, MySQL).
///
/// All methods are `&self` (internal Mutex per driver); `Send + Sync` so the
/// handle can live inside `Arc` inside `Mutex<SqlState>` across threads.
pub(super) trait SqlDriver: Send + Sync {
    fn query(&self, sql: &str, params: &[String]) -> Result<Vec<Value>, String>;
    fn query_columns(&self, sql: &str, params: &[String]) -> Result<Vec<String>, String>;
    fn exec(&self, sql: &str, params: &[String]) -> Result<u64, String>;
    #[allow(dead_code)]
    fn url(&self) -> &str;
    /// Introspection: list all user tables.
    fn tables_sql(&self) -> &'static str;
    /// Introspection: list columns for a given table.
    fn columns_sql(&self, table: &str) -> String;
}

/// Open a driver from a normalised URL (`sqlite:`, `postgres:`, `mysql:`).
pub(super) fn open(url: &str) -> Result<Arc<dyn SqlDriver>, String> {
    // `file:` is SQLite's own URI scheme (`file:app.db?mode=ro`), so it selects
    // the sqlite driver just as `sqlite:` does — the driver passes it through
    // whole and SQLite parses the options.
    if url.starts_with("sqlite:") || url.starts_with("file:") {
        Ok(Arc::new(super::sqlite::SqliteDriver::open(url)?))
    } else if url.starts_with("postgres:") || url.starts_with("postgresql:") {
        Ok(Arc::new(super::postgres::PostgresDriver::open(url)?))
    } else if url.starts_with("mysql:") || url.starts_with("mysql2:") {
        Ok(Arc::new(super::mysql::MySqlDriver::open(url)?))
    } else {
        Err(format!("Unsupported URL scheme: {}", url))
    }
}
