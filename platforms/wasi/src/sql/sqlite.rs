use super::driver::SqlDriver;
use std::sync::{Arc, Mutex};
use vybe_runtime::Value;
use vybe_runtime::value::Object;

pub(super) struct SqliteDriver {
    conn: Mutex<rusqlite::Connection>,
    #[allow(dead_code)]
    url: String,
}

impl SqliteDriver {
    pub(super) fn open(url: &str) -> Result<Self, String> {
        let raw = url.trim_start_matches("sqlite:").trim_start_matches("//");
        // A `file:` URI carries its own options — `?mode=ro`, `?cache=shared`,
        // `?immutable=1` — and `SQLITE_OPEN_URI` below is what tells SQLite to
        // parse them, so the name must reach it WHOLE. Splitting the query off
        // for every name silently dropped every one of those: a caller asking
        // for read-only got a writable database and its writes succeeded.
        //
        // `wasi:sql` only ever says `open(name: string)` and says nothing about
        // the shape of that string (wasi-sql is Phase 1), so how completely we
        // read it is an implementation choice, not a spec question — this adds
        // no interface and no function.
        let path = if raw.starts_with("file:") {
            raw
        } else {
            raw.split('?').next().unwrap_or(raw)
        };
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
        conn.set_prepared_statement_cache_capacity(64);
        Ok(Self {
            conn: Mutex::new(conn),
            url: url.to_string(),
        })
    }
}

fn row_to_obj(
    row: &rusqlite::Row,
    col_names: &[String],
    col_name_values: &[Value],
    index_keys: &[String],
    row_type: &Value,
) -> Value {
    use rusqlite::types::ValueRef;
    let mut obj = super::row_object(col_names.len());
    obj.properties.insert("__type".into(), row_type.clone());
    for (i, name) in col_names.iter().enumerate() {
        let val = match row.get_ref(i).unwrap_or(ValueRef::Null) {
            ValueRef::Null => Value::Null,
            ValueRef::Integer(n) => Value::F64(n as f64),
            ValueRef::Real(f) => Value::F64(f),
            ValueRef::Text(b) => Value::String(Arc::from(std::str::from_utf8(b).unwrap_or(""))),
            ValueRef::Blob(_) => Value::Null,
        };
        obj.properties.insert(name.clone(), val.clone());
        obj.properties.insert(index_keys[i].clone(), val);
    }
    obj.properties.insert(
        "__col_names".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(
            col_name_values.to_vec(),
        ))),
    );
    Value::Object(vybe_runtime::heap::alloc(obj))
}

impl SqlDriver for SqliteDriver {
    fn query(&self, sql: &str, params: &[String]) -> Result<Vec<Value>, String> {
        let conn = self.conn.lock().unwrap();
        let mut stmt = conn.prepare_cached(sql).map_err(|e| e.to_string())?;
        let col_names: Vec<String> = stmt.column_names().iter().map(|s| s.to_string()).collect();
        let col_name_values: Vec<Value> = col_names
            .iter()
            .map(|name| Value::String(Arc::from(name.as_str())))
            .collect();
        let index_keys = super::column_index_keys(col_names.len());
        let row_type = Value::String(Arc::from("DataRow"));
        let rows = stmt
            .query_map(
                rusqlite::params_from_iter(params.iter().map(|s| s.as_str())),
                |row| {
                    Ok(row_to_obj(
                        row,
                        &col_names,
                        &col_name_values,
                        &index_keys,
                        &row_type,
                    ))
                },
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
        conn.prepare_cached(sql)
            .and_then(|mut stmt| {
                stmt.execute(rusqlite::params_from_iter(
                    params.iter().map(|s| s.as_str()),
                ))
            })
            .map(|n| n as u64)
            .map_err(|e| e.to_string())
    }

    fn url(&self) -> &str {
        &self.url
    }

    fn tables_sql(&self) -> &'static str {
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
    }

    fn columns_sql(&self, table: &str) -> String {
        format!("PRAGMA table_info({})", table)
    }
}

#[cfg(test)]
mod cache_tests {
    use super::{SqlDriver, SqliteDriver};
    use vybe_runtime::Value;

    #[test]
    fn repeated_statements_rebind_parameters() {
        let db = SqliteDriver::open("sqlite::memory:").unwrap();
        db.exec("CREATE TABLE items (id INTEGER, label TEXT)", &[])
            .unwrap();
        for id in 0..8 {
            db.exec(
                "INSERT INTO items (id, label) VALUES (?1, ?2)",
                &[id.to_string(), format!("item-{id}")],
            )
            .unwrap();
        }
        assert_eq!(
            db.query_columns("SELECT label FROM items WHERE id = ?1", &[])
                .unwrap(),
            vec!["label"]
        );
        assert_eq!(
            db.query_columns("SELECT * FROM items", &[]).unwrap(),
            vec!["id", "label"]
        );
        for id in (0..8).rev() {
            let rows = db
                .query("SELECT label FROM items WHERE id = ?1", &[id.to_string()])
                .unwrap();
            assert_eq!(rows.len(), 1);
            let Value::Object(row) = &rows[0] else {
                panic!("expected row object")
            };
            let row = row.lock().unwrap();
            assert_eq!(
                row.properties.get("label"),
                Some(&Value::String(format!("item-{id}").into()))
            );
        }
        for id in 0..32 {
            let sql = format!("SELECT ?1 AS value -- statement {id}");
            let rows = db.query(&sql, &[id.to_string()]).unwrap();
            assert_eq!(rows.len(), 1);
            let Value::Object(row) = &rows[0] else {
                panic!("expected row object")
            };
            assert_eq!(
                row.lock().unwrap().properties["value"],
                Value::String(id.to_string().into())
            );
        }
        db.exec("ALTER TABLE items ADD COLUMN note TEXT", &[])
            .unwrap();
        assert_eq!(
            db.query_columns("SELECT * FROM items", &[]).unwrap(),
            vec!["id", "label", "note"]
        );
    }

    #[test]
    fn in_memory_connections_do_not_share_session_state() {
        let first = SqliteDriver::open("sqlite::memory:").unwrap();
        first
            .exec("CREATE TABLE private_rows (value INTEGER)", &[])
            .unwrap();
        let second = SqliteDriver::open("sqlite::memory:").unwrap();
        assert!(second.query("SELECT * FROM private_rows", &[]).is_err());
    }

    #[test]
    #[ignore = "local SQL setup and repeated-query benchmark"]
    fn bench_open_and_repeated_queries() {
        use std::time::Instant;

        let open_start = Instant::now();
        for _ in 0..100 {
            let db = SqliteDriver::open("sqlite::memory:").unwrap();
            drop(db);
        }
        let open_elapsed = open_start.elapsed();

        let db = SqliteDriver::open("sqlite::memory:").unwrap();
        let query_start = Instant::now();
        for id in 0..1000 {
            let rows = db.query("SELECT ?1 AS value", &[id.to_string()]).unwrap();
            assert_eq!(rows.len(), 1);
        }
        eprintln!(
            "SQLite 100 opens: {open_elapsed:?}, 1000 rebound queries: {:?}",
            query_start.elapsed()
        );

        let statements: Vec<String> = (0..32)
            .map(|id| format!("SELECT ?1 AS value -- statement {id}"))
            .collect();
        for capacity in [16, 64] {
            let db = SqliteDriver::open("sqlite::memory:").unwrap();
            db.conn
                .lock()
                .unwrap()
                .set_prepared_statement_cache_capacity(capacity);
            let start = Instant::now();
            for id in 0..3200 {
                let rows = db
                    .query(&statements[id % statements.len()], &[id.to_string()])
                    .unwrap();
                assert_eq!(rows.len(), 1);
            }
            eprintln!(
                "SQLite cache {capacity}, 3200 queries / 32 statements: {:?}",
                start.elapsed()
            );
        }
    }

    #[test]
    #[ignore = "large SQL result materialization benchmark"]
    fn bench_wide_result_materialization() {
        use std::time::Instant;

        let db = SqliteDriver::open("sqlite::memory:").unwrap();
        let sql = "WITH RECURSIVE seq(n) AS (SELECT 0 UNION ALL SELECT n + 1 FROM seq WHERE n < 9999) SELECT n AS c0, n AS c1, n AS c2, n AS c3, n AS c4, n AS c5, n AS c6, n AS c7, n AS c8, n AS c9 FROM seq";
        for _ in 0..3 {
            let start = Instant::now();
            let rows = db.query(sql, &[]).unwrap();
            assert_eq!(rows.len(), 10_000);
            eprintln!("SQLite 10k x 10 rows: {:?}", start.elapsed());
        }
    }
}
