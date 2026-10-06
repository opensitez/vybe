use super::driver::SqlDriver;
use std::borrow::Cow;
use std::collections::{HashMap, VecDeque};
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::Value;
use vybe_runtime::value::Object;

pub(super) struct PostgresDriver {
    session: Mutex<Option<Session>>,
    pool: Arc<Pool>,
    #[allow(dead_code)]
    url: String,
}

impl PostgresDriver {
    pub(super) fn open(url: &str) -> Result<Self, String> {
        let (pool, retired) = {
            let mut pools = pools().lock().unwrap();
            let retired = if !pools.contains_key(url) && pools.len() >= 16 {
                Some(std::mem::take(&mut *pools))
            } else {
                None
            };
            let pool = if let Some(pool) = pools.get(url) {
                pool.clone()
            } else {
                let pool = Arc::new(Pool::new(url));
                pools.insert(url.to_string(), pool.clone());
                pool
            };
            (pool, retired)
        };
        drop(retired);
        let session = pool.checkout()?;
        Ok(Self {
            session: Mutex::new(Some(session)),
            pool,
            url: url.to_string(),
        })
    }
}

struct Session {
    client: postgres::Client,
    prepared: HashMap<String, postgres::Statement>,
    prepared_order: VecDeque<String>,
}

impl Session {
    fn statement(&mut self, sql: &str) -> Result<postgres::Statement, String> {
        if let Some(stmt) = self.prepared.get(sql) {
            return Ok(stmt.clone());
        }
        let stmt = self.client.prepare(sql).map_err(|e| e.to_string())?;
        if self.prepared.len() == 64 {
            if let Some(old) = self.prepared_order.pop_front() {
                self.prepared.remove(&old);
            }
        }
        self.prepared.insert(sql.to_string(), stmt.clone());
        self.prepared_order.push_back(sql.to_string());
        Ok(stmt)
    }

    fn clear_prepared(&mut self) {
        self.prepared.clear();
        self.prepared_order.clear();
    }
}

struct Pool {
    url: String,
    idle: Mutex<Vec<Session>>,
}

impl Pool {
    fn new(url: &str) -> Self {
        Self {
            url: url.to_string(),
            idle: Mutex::new(Vec::new()),
        }
    }

    fn checkout(&self) -> Result<Session, String> {
        if let Some(session) = self.idle.lock().unwrap().pop() {
            return Ok(session);
        }
        let client =
            postgres::Client::connect(&self.url, postgres::NoTls).map_err(|e| e.to_string())?;
        Ok(Session {
            client,
            prepared: HashMap::new(),
            prepared_order: VecDeque::new(),
        })
    }

    fn checkin(&self, mut session: Session) {
        if session.client.is_closed() {
            return;
        }
        // DISCARD ALL resets prepared statements, GUCs and temporary objects.
        // A failed transaction needs ROLLBACK before DISCARD ALL can run.
        if session.client.batch_execute("DISCARD ALL").is_err()
            && (session.client.batch_execute("ROLLBACK").is_err()
                || session.client.batch_execute("DISCARD ALL").is_err())
        {
            return;
        }
        session.clear_prepared();
        let mut idle = self.idle.lock().unwrap();
        if idle.len() < 8 {
            idle.push(session);
        }
    }
}

static POOLS: OnceLock<Mutex<HashMap<String, Arc<Pool>>>> = OnceLock::new();

fn pools() -> &'static Mutex<HashMap<String, Arc<Pool>>> {
    POOLS.get_or_init(|| Mutex::new(HashMap::new()))
}

pub(super) fn reset_pools() {
    if let Some(pools) = POOLS.get() {
        let retired = std::mem::take(&mut *pools.lock().unwrap());
        drop(retired);
    }
}

impl Drop for PostgresDriver {
    fn drop(&mut self) {
        if let Ok(mut session) = self.session.lock() {
            if let Some(session) = session.take() {
                self.pool.checkin(session);
            }
        }
    }
}

/// Rewrite `?` positional placeholders to `$1, $2, …` for PostgreSQL,
/// skipping `?` inside single-quoted string literals.
fn rewrite_placeholders(sql: &str) -> Cow<'_, str> {
    if !sql.as_bytes().contains(&b'?') {
        return Cow::Borrowed(sql);
    }
    let mut out = String::with_capacity(sql.len() + 16);
    let mut n = 1usize;
    let mut in_str = false;
    let mut prev = '\0';
    for ch in sql.chars() {
        if ch == '\'' && prev != '\\' {
            in_str = !in_str;
        }
        if ch == '?' && !in_str {
            out.push_str(&format!("${}", n));
            n += 1;
        } else {
            out.push(ch);
        }
        prev = ch;
    }
    Cow::Owned(out)
}

fn row_to_obj(
    row: &postgres::Row,
    col_names: &[String],
    col_name_values: &[Value],
    index_keys: &[String],
    row_type: &Value,
) -> Value {
    use postgres::types::Type;
    let mut obj = super::row_object(col_names.len());
    obj.properties.insert("__type".into(), row_type.clone());
    for (i, col) in row.columns().iter().enumerate() {
        let val = match *col.type_() {
            Type::BOOL => match row.try_get::<_, Option<bool>>(i) {
                Ok(Some(b)) => Value::Bool(b),
                _ => Value::Null,
            },
            Type::INT2 => row
                .try_get::<_, Option<i16>>(i)
                .ok()
                .flatten()
                .map(|n| Value::F64(n as f64))
                .unwrap_or(Value::Null),
            Type::INT4 => row
                .try_get::<_, Option<i32>>(i)
                .ok()
                .flatten()
                .map(|n| Value::F64(n as f64))
                .unwrap_or(Value::Null),
            Type::INT8 => row
                .try_get::<_, Option<i64>>(i)
                .ok()
                .flatten()
                .map(|n| Value::F64(n as f64))
                .unwrap_or(Value::Null),
            Type::OID => row
                .try_get::<_, Option<u32>>(i)
                .ok()
                .flatten()
                .map(|n| Value::F64(n as f64))
                .unwrap_or(Value::Null),
            Type::FLOAT4 => row
                .try_get::<_, Option<f32>>(i)
                .ok()
                .flatten()
                .map(|n| Value::F64(n as f64))
                .unwrap_or(Value::Null),
            Type::FLOAT8 => match row.try_get::<_, Option<f64>>(i) {
                Ok(Some(f)) => Value::F64(f),
                _ => Value::Null,
            },
            // TEXT, VARCHAR, DATE, TIMESTAMP, UUID, NUMERIC, JSONB, etc.
            _ => match row.try_get::<_, Option<String>>(i) {
                Ok(Some(s)) => super::parse_scalar(&s),
                _ => Value::Null,
            },
        };
        obj.properties.insert(col_names[i].clone(), val.clone());
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

impl SqlDriver for PostgresDriver {
    fn query(&self, sql: &str, params: &[String]) -> Result<Vec<Value>, String> {
        use postgres::fallible_iterator::FallibleIterator;

        let pg_sql = rewrite_placeholders(sql);
        let mut guard = self.session.lock().unwrap();
        let session = guard.as_mut().unwrap();
        let stmt = session.statement(&pg_sql)?;
        let col_names: Vec<String> = stmt
            .columns()
            .iter()
            .map(|col| col.name().to_string())
            .collect();
        let col_name_values: Vec<Value> = col_names
            .iter()
            .map(|name| Value::String(Arc::from(name.as_str())))
            .collect();
        let index_keys = super::column_index_keys(col_names.len());
        let row_type = Value::String(Arc::from("DataRow"));
        let param_refs = params
            .iter()
            .map(|s| s as &(dyn postgres::types::ToSql + Sync));
        let mut result = session
            .client
            .query_raw(&stmt, param_refs)
            .map_err(|error| error.to_string())?;
        let mut rows = Vec::new();
        while let Some(row) = result.next().map_err(|error| error.to_string())? {
            rows.push(row_to_obj(
                &row,
                &col_names,
                &col_name_values,
                &index_keys,
                &row_type,
            ));
        }
        Ok(rows)
    }

    fn query_columns(&self, sql: &str, _params: &[String]) -> Result<Vec<String>, String> {
        let pg_sql = rewrite_placeholders(sql);
        let mut guard = self.session.lock().unwrap();
        let session = guard.as_mut().unwrap();
        let stmt = session.statement(&pg_sql)?;
        Ok(stmt
            .columns()
            .iter()
            .map(|col| col.name().to_string())
            .collect())
    }

    fn exec(&self, sql: &str, params: &[String]) -> Result<u64, String> {
        let pg_sql = rewrite_placeholders(sql);
        let mut guard = self.session.lock().unwrap();
        let session = guard.as_mut().unwrap();
        let param_refs: Vec<&(dyn postgres::types::ToSql + Sync)> = params
            .iter()
            .map(|s| s as &(dyn postgres::types::ToSql + Sync))
            .collect();
        let command = pg_sql.trim_start().split_whitespace().next().unwrap_or("");
        if matches!(
            command.to_ascii_uppercase().as_str(),
            "SELECT" | "INSERT" | "UPDATE" | "DELETE" | "WITH"
        ) {
            let stmt = session.statement(&pg_sql)?;
            session
                .client
                .execute(&stmt, &param_refs)
                .map_err(|e| e.to_string())
        } else {
            session.clear_prepared();
            session
                .client
                .execute(pg_sql.as_ref(), &param_refs)
                .map_err(|e| e.to_string())
        }
    }

    fn url(&self) -> &str {
        &self.url
    }

    fn tables_sql(&self) -> &'static str {
        "SELECT table_name AS name FROM information_schema.tables \
         WHERE table_schema = 'public' ORDER BY table_name"
    }

    fn columns_sql(&self, table: &str) -> String {
        format!(
            "SELECT column_name FROM information_schema.columns \
             WHERE table_name = '{}' ORDER BY ordinal_position",
            table
        )
    }
}

#[cfg(test)]
mod pool_tests {
    use super::{PostgresDriver, SqlDriver, rewrite_placeholders};
    use std::borrow::Cow;
    use vybe_runtime::Value;

    fn first_cell(rows: Vec<Value>, column: &str) -> Value {
        let Value::Object(row) = &rows[0] else {
            panic!("expected SQL row")
        };
        row.lock().unwrap().properties[column].clone()
    }

    #[test]
    fn unbound_sql_uses_borrowed_text() {
        let sql = "SELECT 1 AS value";
        assert!(matches!(rewrite_placeholders(sql), Cow::Borrowed(text) if text == sql));
        assert_eq!(
            rewrite_placeholders("SELECT ? AS value"),
            "SELECT $1 AS value"
        );
    }

    #[test]
    fn reuses_clean_connection_and_prepared_queries() {
        let Ok(url) = std::env::var("VYBE_SQL_TEST_POSTGRES_URL") else {
            return;
        };
        let first = PostgresDriver::open(&url).unwrap();
        let pid = first_cell(
            first.query("SELECT pg_backend_pid() AS pid", &[]).unwrap(),
            "pid",
        );
        let initial_name = first_cell(
            first.query("SHOW application_name", &[]).unwrap(),
            "application_name",
        );
        let parallel = PostgresDriver::open(&url).unwrap();
        assert_ne!(
            first_cell(
                parallel
                    .query("SELECT pg_backend_pid() AS pid", &[])
                    .unwrap(),
                "pid"
            ),
            pid,
        );
        drop(parallel);
        for input in ["alpha", "beta", "gamma"] {
            assert_eq!(
                first_cell(
                    first
                        .query("SELECT ?::text AS value", &[input.into()])
                        .unwrap(),
                    "value"
                ),
                Value::String(input.into()),
            );
        }
        for _ in 0..2 {
            assert_eq!(
                first
                    .query_columns("SELECT ?::text AS label", &[])
                    .unwrap(),
                vec!["label"],
            );
        }
        assert_eq!(
            first
                .session
                .lock()
                .unwrap()
                .as_ref()
                .unwrap()
                .prepared
                .len(),
            4
        );
        first
            .exec("SET application_name = 'vybe_pool_marker'", &[])
            .unwrap();
        first.exec("BEGIN", &[]).unwrap();
        drop(first);

        let second = PostgresDriver::open(&url).unwrap();
        assert_eq!(
            first_cell(
                second.query("SELECT pg_backend_pid() AS pid", &[]).unwrap(),
                "pid"
            ),
            pid
        );
        assert_eq!(
            first_cell(
                second.query("SHOW application_name", &[]).unwrap(),
                "application_name"
            ),
            initial_name,
        );
        second.exec("BEGIN", &[]).unwrap();
        assert!(second.query("SELECT 1 / 0 AS broken", &[]).is_err());
        drop(second);

        let recovered = PostgresDriver::open(&url).unwrap();
        assert_eq!(
            first_cell(
                recovered
                    .query("SELECT pg_backend_pid() AS pid", &[])
                    .unwrap(),
                "pid"
            ),
            pid,
        );
        assert_eq!(
            first_cell(recovered.query("SELECT 1 AS value", &[]).unwrap(), "value"),
            Value::F64(1.0),
        );
    }
}
