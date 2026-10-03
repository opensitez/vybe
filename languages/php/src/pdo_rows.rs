//! Convert shared SQL DataRows to the representation requested by PHP PDO.
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Framework, Value, heap};

fn fetch_mode(value: Option<&Value>) -> u32 {
    match value {
        Some(Value::BigInt(number)) => number.to_f64() as u32,
        Some(value) => value.as_f64() as u32,
        None => 0,
    }
}

fn fetch_row(row: &Value, mode: u32) -> Value {
    let mode = match mode & 0xffff {
        0 => 4,
        other => other,
    };
    if !matches!(mode, 2 | 3 | 4 | 5) {
        return row.clone();
    }
    let Value::Object(source) = row else {
        return row.clone();
    };
    let source = source.lock().unwrap();
    let Some(Value::Object(names)) = source.properties.get("__col_names") else {
        return row.clone();
    };
    let names = names.lock().unwrap();
    let ObjectKind::Array(names) = &names.kind else {
        return row.clone();
    };
    let mut output = Object::new();
    if mode == 5 {
        for name in names {
            if let Value::String(name) = name {
                output.properties.insert(
                    name.to_string(),
                    source
                        .properties
                        .get(name.as_ref())
                        .cloned()
                        .unwrap_or(Value::Null),
                );
            }
        }
    } else {
        output.kind = ObjectKind::Map(Default::default());
        let mut values = Vec::with_capacity(names.len());
        for (index, name) in names.iter().enumerate() {
            let value = source
                .properties
                .get(&index.to_string())
                .cloned()
                .unwrap_or(Value::Null);
            if mode == 3 {
                values.push(value);
                continue;
            }
            if let Value::String(name) = name {
                let numeric = name
                    .parse::<i64>()
                    .ok()
                    .filter(|number| number.to_string() == name.as_ref());
                let key = numeric
                    .map(Value::I64)
                    .unwrap_or_else(|| Value::String(name.clone()));
                // Association order follows the driver's column metadata.
                if let ObjectKind::Map(ref mut map) = output.kind {
                    map.insert(key, value.clone());
                }
            }
            if mode == 4 {
                if let ObjectKind::Map(ref mut map) = output.kind {
                    map.insert(Value::I64(index as i64), value);
                }
            }
        }
        if mode == 3 {
            output.kind = ObjectKind::Array(values);
        }
    }
    Value::Object(heap::alloc(output))
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:pdo",
        "fetchRow",
        Box::new(|_, args| {
            fetch_row(
                args.first().unwrap_or(&Value::Null),
                fetch_mode(args.get(1)),
            )
        }),
    );
    fw.register_host_fn(
        "php:pdo",
        "fetchRows",
        Box::new(|_, args| {
            let Some(Value::Object(rows)) = args.first() else {
                return Value::Null;
            };
            let rows = rows.lock().unwrap();
            let ObjectKind::Array(rows) = &rows.kind else {
                return args[0].clone();
            };
            let mode = fetch_mode(args.get(1));
            Value::Object(heap::alloc(Object::new_array(
                rows.iter().map(|row| fetch_row(row, mode)).collect(),
            )))
        }),
    );
}
