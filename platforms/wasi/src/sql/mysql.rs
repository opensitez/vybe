use super::driver::SqlDriver;
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
        let conn = mysql::Conn::new(opts).map_err(|e| e.to_string())?;
        Ok(Self {
            conn: Mutex::new(conn),
            url,
        })
    }
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
            Some(Ok(mysql::Value::Date(y, mo, d, h, mi, s, _))) => Value::String(Arc::from(
                format!("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", y, mo, d, h, mi, s).as_str(),
            )),
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

impl SqlDriver for MySqlDriver {
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
