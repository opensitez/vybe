use super::driver::SqlDriver;
use std::path::Path;
use std::sync::{Arc, Mutex};
use vybe_runtime::Value;
use vybe_runtime::value::Object;

pub(super) struct MySqlDriver {
    conn: Mutex<mysql::Conn>,
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
        let opts = mysql::Opts::from_url(&url).map_err(|e| e.to_string())?;
        if opts.addr_is_loopback() && opts.get_socket().is_none() {
            if let Ok(conn) = open_loopback_socket(&opts) {
                return Ok(Self {
                    conn: Mutex::new(conn),
                    url,
                });
            }
        }
        let conn = match mysql::Conn::new(opts.clone()) {
            Ok(conn) => conn,
            Err(first_err) => {
                if opts.addr_is_loopback() && opts.get_socket().is_none() {
                    open_loopback_socket(&opts).map_err(|_| first_err.to_string())?
                } else {
                    return Err(first_err.to_string());
                }
            }
        };
        Ok(Self {
            conn: Mutex::new(conn),
            url,
        })
    }
}

fn open_loopback_socket(opts: &mysql::Opts) -> Result<mysql::Conn, String> {
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

    let mut last = None;
    for socket in candidates {
        if !Path::new(&socket).exists() {
            continue;
        }
        let socket_opts: mysql::Opts = mysql::OptsBuilder::from_opts(opts.clone())
            .socket(Some(socket))
            .into();
        match mysql::Conn::new(socket_opts) {
            Ok(conn) => return Ok(conn),
            Err(err) => last = Some(err.to_string()),
        }
    }
    Err(last.unwrap_or_else(|| "no MySQL Unix socket candidate was available".to_string()))
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

fn row_to_obj(row: &mysql::Row) -> Value {
    let col_names: Vec<String> = row
        .columns_ref()
        .iter()
        .map(|c| c.name_str().to_string())
        .collect();
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("DataRow")));
    for (i, name) in col_names.iter().enumerate() {
        let val = match row.get_opt::<mysql::Value, _>(i) {
            Some(Ok(mysql::Value::NULL)) | None => Value::Null,
            Some(Ok(mysql::Value::Bytes(b))) => super::parse_scalar(&String::from_utf8_lossy(&b)),
            Some(Ok(mysql::Value::Int(n))) => Value::F64(n as f64),
            Some(Ok(mysql::Value::UInt(n))) => Value::F64(n as f64),
            Some(Ok(mysql::Value::Float(f))) => Value::F64(f as f64),
            Some(Ok(mysql::Value::Double(f))) => Value::F64(f),
            Some(Ok(mysql::Value::Date(y, mo, d, h, mi, s, _))) => {
                Value::String(Arc::from(format_date_value(
                    row.columns_ref()[i].column_type(), y, mo, d, h, mi, s,
                )))
            }
            Some(Ok(mysql::Value::Time(neg, days, h, mi, s, _))) => {
                let sign = if neg { "-" } else { "" };
                Value::String(Arc::from(
                    format!("{}{:02}:{:02}:{:02}", sign, days * 24 + h as u32, mi, s).as_str(),
                ))
            }
            _ => Value::Null,
        };
        obj.properties.insert(name.clone(), val.clone());
        obj.properties.insert(i.to_string(), val);
    }
    obj.properties.insert(
        "__col_names".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(
            col_names
                .iter()
                .map(|name| Value::String(Arc::from(name.as_str())))
                .collect(),
        ))),
    );
    Value::Object(vybe_runtime::heap::alloc(obj))
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
    if matches!(column_type, ColumnType::MYSQL_TYPE_DATE | ColumnType::MYSQL_TYPE_NEWDATE) {
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
        assert_eq!(format_date_value(ColumnType::MYSQL_TYPE_DATE, 981, 1, 1, 0, 0, 0), "0981-01-01");
        assert_eq!(format_date_value(ColumnType::MYSQL_TYPE_NEWDATE, 2025, 10, 31, 0, 0, 0), "2025-10-31");
    }

    #[test]
    fn midnight_datetime_is_not_a_date_column() {
        assert_eq!(format_date_value(ColumnType::MYSQL_TYPE_DATETIME, 2025, 10, 31, 0, 0, 0), "2025-10-31 00:00:00");
        assert_eq!(format_date_value(ColumnType::MYSQL_TYPE_TIMESTAMP, 2026, 10, 3, 15, 39, 22), "2026-10-03 15:39:22");
    }
}

impl SqlDriver for MySqlDriver {
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
        let mysql_params: Vec<mysql::Value> = params.iter().map(|s| to_param(s)).collect();
        let mut conn = self.conn.lock().unwrap();
        conn.exec::<mysql::Row, _, _>(sql, mysql::Params::Positional(mysql_params))
            .map(|rows| rows.iter().map(row_to_obj).collect())
            .map_err(|e| e.to_string())
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
        let mysql_params: Vec<mysql::Value> = params.iter().map(|s| to_param(s)).collect();
        let mut conn = self.conn.lock().unwrap();
        conn.exec_drop(sql, mysql::Params::Positional(mysql_params))
            .map_err(|e| e.to_string())?;
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
