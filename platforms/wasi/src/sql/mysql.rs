use super::driver::SqlDriver;
use std::collections::HashMap;
use std::path::Path;
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::Value;
use vybe_runtime::value::Object;

pub(super) struct MySqlDriver {
    conn: Mutex<mysql::PooledConn>,
    #[allow(dead_code)]
    url: String,
}

impl MySqlDriver {
    pub(super) fn open(url: &str) -> Result<Self, String> {
        // Normalise mysql2:// (Ruby/PHP convention) → mysql://
        let url = if url.starts_with("mysql2:") {
            format!("mysql:{}", &url["mysql2:".len()..])
        } else {
            url.to_string()
        };
        let slot = pool_slot(&url);
        let mut initialized = slot.lock().unwrap();
        let conn = if let Some(pool) = initialized.as_ref() {
            let pool = pool.clone();
            drop(initialized);
            pool.get_conn().map_err(|e| e.to_string())?
        } else {
            let opts = mysql::Opts::from_url(&url).map_err(|e| e.to_string())?;
            let (pool, conn) = open_pool(&opts)?;
            *initialized = Some(pool);
            conn
        };
        Ok(Self {
            conn: Mutex::new(conn),
            url,
        })
    }
}

type PoolSlot = Arc<Mutex<Option<mysql::Pool>>>;
static POOLS: OnceLock<Mutex<HashMap<String, PoolSlot>>> = OnceLock::new();

fn pools() -> &'static Mutex<HashMap<String, PoolSlot>> {
    POOLS.get_or_init(|| Mutex::new(HashMap::new()))
}

fn pool_slot(url: &str) -> PoolSlot {
    let (slot, retired) = {
        let mut pools = pools().lock().unwrap();
        if let Some(slot) = pools.get(url) {
            return slot.clone();
        }
        let retired = if pools.len() >= 16 {
            Some(std::mem::take(&mut *pools))
        } else {
            None
        };
        let slot = Arc::new(Mutex::new(None));
        pools.insert(url.to_string(), slot.clone());
        (slot, retired)
    };
    drop(retired);
    slot
}

pub(super) fn reset_pools() {
    if let Some(pools) = POOLS.get() {
        let retired = std::mem::take(&mut *pools.lock().unwrap());
        drop(retired);
    }
}

fn open_pool(opts: &mysql::Opts) -> Result<(mysql::Pool, mysql::PooledConn), String> {
    let pool_opts = mysql::PoolOpts::default()
        .with_constraints(mysql::PoolConstraints::new_const::<0, 64>())
        .with_reset_connection(true);
    let mut candidates = Vec::new();
    if opts.addr_is_loopback() && opts.get_socket().is_none() {
        for socket in loopback_sockets() {
            if Path::new(&socket).exists() {
                candidates.push(
                    mysql::OptsBuilder::from_opts(opts.clone())
                        .socket(Some(socket))
                        .into(),
                );
            }
        }
    }
    candidates.push(opts.clone());
    let mut last_error = None;
    for candidate in candidates {
        let configured: mysql::Opts = mysql::OptsBuilder::from_opts(candidate)
            .pool_opts(Some(pool_opts.clone()))
            .into();
        match mysql::Pool::new(configured) {
            Ok(pool) => match pool.get_conn() {
                Ok(conn) => return Ok((pool, conn)),
                Err(err) => last_error = Some(err.to_string()),
            },
            Err(err) => last_error = Some(err.to_string()),
        }
    }
    Err(last_error.unwrap_or_else(|| "could not open MySQL connection".to_string()))
}

fn loopback_sockets() -> Vec<String> {
    let mut candidates = Vec::new();
    if let Ok(socket) = std::env::var("MYSQL_UNIX_PORT") {
        candidates.push(socket);
    }
    candidates.extend(
        [
            "/tmp/mysql.sock",
            "/var/run/mysqld/mysqld.sock",
            "/var/lib/mysql/mysql.sock",
            "/opt/homebrew/var/mysql/mysql.sock",
            "/usr/local/var/mysql/mysql.sock",
            "/Applications/MAMP/tmp/mysql/mysql.sock",
        ]
        .iter()
        .map(|path| path.to_string()),
    );

    candidates
}

fn to_param(s: &str) -> mysql::Value {
    if let Ok(n) = s.parse::<i64>() {
        return mysql::Value::Int(n);
    }
    if let Ok(f) = s.parse::<f64>() {
        return mysql::Value::Double(f);
    }
    mysql::Value::Bytes(s.as_bytes().to_vec())
}

fn row_to_obj(
    row: mysql::Row,
    col_names: &[String],
    col_types: &[mysql::consts::ColumnType],
    col_name_values: &[Value],
    index_keys: &[String],
    row_type: &Value,
) -> Value {
    let mut obj = super::row_object(col_names.len());
    obj.properties.insert("__type".into(), row_type.clone());
    for (i, (name, cell)) in col_names.iter().zip(row.unwrap_raw()).enumerate() {
        let val = match cell {
            Some(mysql::Value::NULL) | None => Value::Null,
            Some(mysql::Value::Bytes(b)) => super::parse_scalar(&String::from_utf8_lossy(&b)),
            Some(mysql::Value::Int(n)) => Value::F64(n as f64),
            Some(mysql::Value::UInt(n)) => Value::F64(n as f64),
            Some(mysql::Value::Float(f)) => Value::F64(f as f64),
            Some(mysql::Value::Double(f)) => Value::F64(f),
            Some(mysql::Value::Date(y, mo, d, h, mi, s, _)) => Value::String(Arc::from(
                format_date_value(col_types[i], y, mo, d, h, mi, s),
            )),
            Some(mysql::Value::Time(neg, days, h, mi, s, _)) => {
                let sign = if neg { "-" } else { "" };
                Value::String(Arc::from(
                    format!("{}{:02}:{:02}:{:02}", sign, days * 24 + h as u32, mi, s).as_str(),
                ))
            }
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

fn collect_rows<T: mysql::prelude::Protocol>(
    result: mysql::QueryResult<'_, '_, '_, T>,
) -> Result<Vec<Value>, String> {
    collect_rows_with_metadata(result, false).map(|(rows, _)| rows)
}

fn collect_rows_with_metadata<T: mysql::prelude::Protocol>(
    mut result: mysql::QueryResult<'_, '_, '_, T>,
    include_metadata: bool,
) -> Result<(Vec<Value>, Vec<Value>), String> {
    let columns = result.columns();
    let metadata = if include_metadata {
        columns.as_ref().iter().map(|column| {
            let mut descriptor = Object::new();
            for (key, value) in [
                ("name", column.name_str().into_owned()),
                ("orgname", column.org_name_str().into_owned()),
                ("table", column.table_str().into_owned()),
                ("orgtable", column.org_table_str().into_owned()),
                ("db", column.schema_str().into_owned()),
            ] {
                descriptor.properties.insert(key.into(), Value::String(Arc::from(value)));
            }
            for (key, value) in [
                ("type", column.column_type() as u8 as f64),
                ("flags", column.flags().bits() as f64),
                ("charsetnr", column.character_set() as f64),
                ("length", column.column_length() as f64),
                ("decimals", column.decimals() as f64),
            ] {
                descriptor.properties.insert(key.into(), Value::F64(value));
            }
            Value::Object(vybe_runtime::heap::alloc(descriptor))
        }).collect()
    } else { Vec::new() };
    let col_names: Vec<String> = columns
        .as_ref()
        .iter()
        .map(|col| col.name_str().to_string())
        .collect();
    let col_types: Vec<mysql::consts::ColumnType> = columns
        .as_ref()
        .iter()
        .map(|col| col.column_type())
        .collect();
    let col_name_values: Vec<Value> = col_names
        .iter()
        .map(|name| Value::String(Arc::from(name.as_str())))
        .collect();
    let index_keys = super::column_index_keys(col_names.len());
    let row_type = Value::String(Arc::from("DataRow"));

    let mut rows = Vec::new();
    if let Some(first_set) = result.iter() {
        for row in first_set {
            rows.push(row_to_obj(
                row.map_err(|error| error.to_string())?,
                &col_names,
                &col_types,
                &col_name_values,
                &index_keys,
                &row_type,
            ));
        }
    }
    Ok((rows, metadata))
}

fn format_date_value(
    column_type: mysql::consts::ColumnType,
    year: u16,
    month: u8,
    day: u8,
    hour: u8,
    minute: u8,
    second: u8,
) -> String {
    use mysql::consts::ColumnType;
    if matches!(
        column_type,
        ColumnType::MYSQL_TYPE_DATE | ColumnType::MYSQL_TYPE_NEWDATE
    ) {
        format!("{year:04}-{month:02}-{day:02}")
    } else {
        format!("{year:04}-{month:02}-{day:02} {hour:02}:{minute:02}:{second:02}")
    }
}

#[cfg(test)]
mod date_value_tests {
    use super::format_date_value;
    use mysql::consts::ColumnType;

    #[test]
    fn date_columns_keep_date_only_text() {
        assert_eq!(
            format_date_value(ColumnType::MYSQL_TYPE_DATE, 981, 1, 1, 0, 0, 0),
            "0981-01-01"
        );
        assert_eq!(
            format_date_value(ColumnType::MYSQL_TYPE_NEWDATE, 2025, 10, 31, 0, 0, 0),
            "2025-10-31"
        );
    }

    #[test]
    fn midnight_datetime_is_not_a_date_column() {
        assert_eq!(
            format_date_value(ColumnType::MYSQL_TYPE_DATETIME, 2025, 10, 31, 0, 0, 0),
            "2025-10-31 00:00:00"
        );
        assert_eq!(
            format_date_value(ColumnType::MYSQL_TYPE_TIMESTAMP, 2026, 10, 3, 15, 39, 22),
            "2026-10-03 15:39:22"
        );
    }
}

impl SqlDriver for MySqlDriver {
    fn query_with_metadata(&self, sql: &str, params: &[String])
        -> Result<(Vec<Value>, Vec<Value>), String>
    {
        use mysql::prelude::Queryable;
        let mut conn = self.conn.lock().unwrap();
        if params.is_empty() {
            let result = conn.query_iter(sql).map_err(|error| error.to_string())?;
            collect_rows_with_metadata(result, true)
        } else {
            let mysql_params: Vec<mysql::Value> = params.iter().map(|s| to_param(s)).collect();
            let result = conn.exec_iter(sql, mysql::Params::Positional(mysql_params))
                .map_err(|error| error.to_string())?;
            collect_rows_with_metadata(result, true)
        }
    }
    fn select_database(&self, database: &str) -> Result<(), String> {
        use mysql::prelude::Queryable;
        if database.is_empty() || database.contains('\0') {
            return Err("invalid MySQL database name".to_string());
        }
        let quoted = database.replace('`', "``");
        self.conn
            .lock()
            .unwrap()
            .query_drop(format!("USE `{quoted}`"))
            .map_err(|e| e.to_string())
    }

    fn query(&self, sql: &str, params: &[String]) -> Result<Vec<Value>, String> {
        use mysql::prelude::Queryable;
        let mut conn = self.conn.lock().unwrap();
        if params.is_empty() {
            let result = conn.query_iter(sql).map_err(|error| error.to_string())?;
            collect_rows(result)
        } else {
            let mysql_params: Vec<mysql::Value> = params.iter().map(|s| to_param(s)).collect();
            let result = conn
                .exec_iter(sql, mysql::Params::Positional(mysql_params))
                .map_err(|error| error.to_string())?;
            collect_rows(result)
        }
    }

    fn query_columns(&self, sql: &str, _params: &[String]) -> Result<Vec<String>, String> {
        use mysql::prelude::Queryable;
        let mut conn = self.conn.lock().unwrap();
        let stmt = conn.prep(sql).map_err(|e| e.to_string())?;
        Ok(stmt
            .columns()
            .iter()
            .map(|col| col.name_str().to_string())
            .collect())
    }

    fn exec(&self, sql: &str, params: &[String]) -> Result<u64, String> {
        use mysql::prelude::Queryable;
        let mut conn = self.conn.lock().unwrap();
        if params.is_empty() {
            conn.query_drop(sql).map_err(|e| e.to_string())?;
        } else {
            let mysql_params: Vec<mysql::Value> = params.iter().map(|s| to_param(s)).collect();
            conn.exec_drop(sql, mysql::Params::Positional(mysql_params))
                .map_err(|e| e.to_string())?;
        }
        Ok(conn.affected_rows())
    }

    fn url(&self) -> &str {
        &self.url
    }

    fn tables_sql(&self) -> &'static str {
        "SELECT table_name AS name FROM information_schema.tables \
         WHERE table_schema = DATABASE() ORDER BY table_name"
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
    use super::{MySqlDriver, SqlDriver, pool_slot, pools};
    use std::sync::Arc;
    use vybe_runtime::Value;

    #[test]
    fn concurrent_cold_opens_share_one_pool_slot() {
        let url = format!("mysql://pool-slot-test/{}", super::super::state::next_id());
        let slots = std::thread::scope(|scope| {
            let handles: Vec<_> = (0..16).map(|_| scope.spawn(|| pool_slot(&url))).collect();
            handles.into_iter().map(|handle| handle.join().unwrap()).collect::<Vec<_>>()
        });
        assert!(slots.iter().all(|slot| Arc::ptr_eq(slot, &slots[0])));
        let retired = pools().lock().unwrap().remove(&url);
        drop(retired);
    }

    fn first_cell(rows: Vec<Value>, column: &str) -> Value {
        let Value::Object(row) = &rows[0] else {
            panic!("expected SQL row")
        };
        row.lock().unwrap().properties[column].clone()
    }

    #[test]
    fn reuses_connection_without_leaking_session_state() {
        let Ok(url) = std::env::var("VYBE_SQL_TEST_MYSQL_URL") else {
            return;
        };
        let first = MySqlDriver::open(&url).unwrap();
        let id = first_cell(
            first.query("SELECT CONNECTION_ID() AS id", &[]).unwrap(),
            "id",
        );
        let initial_database =
            first_cell(first.query("SELECT DATABASE() AS db", &[]).unwrap(), "db");
        let parallel = MySqlDriver::open(&url).unwrap();
        assert_ne!(
            first_cell(
                parallel.query("SELECT CONNECTION_ID() AS id", &[]).unwrap(),
                "id"
            ),
            id,
        );
        first.exec("START TRANSACTION", &[]).unwrap();
        first.exec("SET @vybe_pool_marker = 7", &[]).unwrap();
        drop(first);
        drop(parallel);

        let second = MySqlDriver::open(&url).unwrap();
        assert_eq!(
            first_cell(
                second.query("SELECT CONNECTION_ID() AS id", &[]).unwrap(),
                "id"
            ),
            id
        );
        assert_eq!(
            first_cell(
                second
                    .query("SELECT @vybe_pool_marker AS marker", &[])
                    .unwrap(),
                "marker"
            ),
            Value::Null
        );
        assert_eq!(
            first_cell(second.query("SELECT DATABASE() AS db", &[]).unwrap(), "db"),
            initial_database
        );
        let payload = first_cell(
            second
                .query("SELECT REPEAT('x', 65536) AS payload", &[])
                .unwrap(),
            "payload",
        );
        let Value::String(payload) = payload else {
            panic!("expected text payload")
        };
        assert_eq!(payload.len(), 65536);
        assert!(payload.bytes().all(|byte| byte == b'x'));
        assert_eq!(
            first_cell(
                second
                    .query("SELECT CAST('2024-02-29' AS DATE) AS date_value", &[])
                    .unwrap(),
                "date_value",
            ),
            Value::String("2024-02-29".into()),
        );
        assert_eq!(
            first_cell(
                second
                    .query("SELECT ? AS bound_value", &["42".into()])
                    .unwrap(),
                "bound_value",
            ),
            Value::F64(42.0),
        );
    }
}
