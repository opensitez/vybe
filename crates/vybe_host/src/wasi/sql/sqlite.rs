use std::sync::{Arc, Mutex};
use vybe_bytecode::Value;
use vybe_bytecode::value::Object;
use super::driver::SqlDriver;

pub(super) struct SqliteDriver {
    conn: Mutex<rusqlite::Connection>,
    url: String,
}

impl SqliteDriver {
    pub(super) fn open(url: &str) -> Result<Self, String> {
        let raw = url.trim_start_matches("sqlite:").trim_start_matches("//");
        let path = raw.split('?').next().unwrap_or(raw);
        let conn = if path == ":memory:" {
            rusqlite::Connection::open_in_memory()
        } else {
            rusqlite::Connection::open_with_flags(
                path,
                rusqlite::OpenFlags::SQLITE_OPEN_READ_WRITE
                    | rusqlite::OpenFlags::SQLITE_OPEN_CREATE
                    | rusqlite::OpenFlags::SQLITE_OPEN_URI,
            )
        }
        .map_err(|e| e.to_string())?;
        Ok(Self { conn: Mutex::new(conn), url: url.to_string() })
    }
}

fn row_to_obj(row: &rusqlite::Row, col_names: &[String]) -> Value {
    use rusqlite::types::ValueRef;
    let mut obj = Object::new();
    obj.properties.insert("__type".into(), Value::String(Arc::from("DataRow")));
    for (i, name) in col_names.iter().enumerate() {
        let val = match row.get_ref(i).unwrap_or(ValueRef::Null) {
            ValueRef::Null       => Value::Null,
            ValueRef::Integer(n) => Value::F64(n as f64),
            ValueRef::Real(f)    => Value::F64(f),
            ValueRef::Text(b)    => Value::String(Arc::from(std::str::from_utf8(b).unwrap_or(""))),
            ValueRef::Blob(_)    => Value::Null,
        };
            obj.properties.insert(name.clone(), val.clone());
            obj.properties.insert(i.to_string(), val);
    }
    obj.properties.insert(
        "__col_names".into(),
        Value::Object(Arc::new(Mutex::new(Object::new_array(
            col_names.iter().map(|name| Value::String(Arc::from(name.as_str()))).collect(),
        )))),
    );
    Value::Object(Arc::new(Mutex::new(obj)))
}

impl SqlDriver for SqliteDriver {
    fn query(&self, sql: &str, params: &[String]) -> Result<Vec<Value>, String> {
        let conn = self.conn.lock().unwrap();
        let mut stmt = conn.prepare(sql).map_err(|e| e.to_string())?;
        let col_names: Vec<String> = stmt.column_names().iter().map(|s| s.to_string()).collect();
        let rows = stmt
            .query_map(
                rusqlite::params_from_iter(params.iter().map(|s| s.as_str())),
                |row| Ok(row_to_obj(row, &col_names)),
            )
            .map_err(|e| e.to_string())?
            .filter_map(|r| r.ok())
            .collect();
        Ok(rows)
    }

    fn query_columns(&self, sql: &str, _params: &[String]) -> Result<Vec<String>, String> {
        let conn = self.conn.lock().unwrap();
        let stmt = conn.prepare(sql).map_err(|e| e.to_string())?;
        Ok(stmt.column_names().iter().map(|s| s.to_string()).collect())
    }

    fn exec(&self, sql: &str, params: &[String]) -> Result<u64, String> {
        let conn = self.conn.lock().unwrap();
        conn.execute(
            sql,
            rusqlite::params_from_iter(params.iter().map(|s| s.as_str())),
        )
        .map(|n| n as u64)
        .map_err(|e| e.to_string())
    }

    fn url(&self) -> &str { &self.url }

    fn tables_sql(&self) -> &'static str {
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
    }

    fn columns_sql(&self, table: &str) -> String {
        format!("PRAGMA table_info({})", table)
    }
}
