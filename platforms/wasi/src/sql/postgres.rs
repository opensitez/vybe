use super::driver::SqlDriver;
use std::sync::{Arc, Mutex};
use vybe_runtime::Value;
use vybe_runtime::value::Object;

pub(super) struct PostgresDriver {
    client: Mutex<postgres::Client>,
    #[allow(dead_code)]
    url: String,
}

impl PostgresDriver {
    pub(super) fn open(url: &str) -> Result<Self, String> {
        let client = postgres::Client::connect(url, postgres::NoTls).map_err(|e| e.to_string())?;
        Ok(Self {
            client: Mutex::new(client),
            url: url.to_string(),
        })
    }
}

/// Rewrite `?` positional placeholders to `$1, $2, …` for PostgreSQL,
/// skipping `?` inside single-quoted string literals.
fn rewrite_placeholders(sql: &str) -> String {
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
    out
}

fn row_to_obj(row: &postgres::Row) -> Value {
    use postgres::types::Type;
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("DataRow")));
    let mut col_names = Vec::new();
    for (i, col) in row.columns().iter().enumerate() {
        let name = col.name().to_string();
        col_names.push(name.clone());
        let val = match *col.type_() {
            Type::BOOL => match row.try_get::<_, Option<bool>>(i) {
                Ok(Some(b)) => Value::Bool(b),
                _ => Value::Null,
            },
            Type::INT2 | Type::INT4 | Type::INT8 | Type::OID => {
                match row.try_get::<_, Option<i64>>(i) {
                    Ok(Some(n)) => Value::F64(n as f64),
                    _ => Value::Null,
                }
            }
            Type::FLOAT4 | Type::FLOAT8 => match row.try_get::<_, Option<f64>>(i) {
                Ok(Some(f)) => Value::F64(f),
                _ => Value::Null,
            },
            // TEXT, VARCHAR, DATE, TIMESTAMP, UUID, NUMERIC, JSONB, etc.
            _ => match row.try_get::<_, Option<String>>(i) {
                Ok(Some(s)) => super::parse_scalar(&s),
                _ => Value::Null,
            },
        };
        obj.properties.insert(name, val.clone());
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

impl SqlDriver for PostgresDriver {
    fn query(&self, sql: &str, params: &[String]) -> Result<Vec<Value>, String> {
        let pg_sql = rewrite_placeholders(sql);
        let mut client = self.client.lock().unwrap();
        let param_refs: Vec<&(dyn postgres::types::ToSql + Sync)> = params
            .iter()
            .map(|s| s as &(dyn postgres::types::ToSql + Sync))
            .collect();
        client
            .query(pg_sql.as_str(), &param_refs)
            .map(|rows| rows.iter().map(row_to_obj).collect())
            .map_err(|e| e.to_string())
    }

    fn query_columns(&self, sql: &str, _params: &[String]) -> Result<Vec<String>, String> {
        let pg_sql = rewrite_placeholders(sql);
        let mut client = self.client.lock().unwrap();
        let stmt = client.prepare(pg_sql.as_str()).map_err(|e| e.to_string())?;
        Ok(stmt
            .columns()
            .iter()
            .map(|col| col.name().to_string())
            .collect())
    }

    fn exec(&self, sql: &str, params: &[String]) -> Result<u64, String> {
        let pg_sql = rewrite_placeholders(sql);
        let mut client = self.client.lock().unwrap();
        let param_refs: Vec<&(dyn postgres::types::ToSql + Sync)> = params
            .iter()
            .map(|s| s as &(dyn postgres::types::ToSql + Sync))
            .collect();
        client
            .execute(pg_sql.as_str(), &param_refs)
            .map_err(|e| e.to_string())
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
