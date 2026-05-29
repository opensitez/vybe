//! System.Data — DataTable, DataSet (simplified objects)

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

fn string_value(value: &str) -> Value {
    Value::String(Arc::from(value))
}

fn row_column_names(row: &Arc<Mutex<Object>>) -> Vec<String> {
    let guard = row.lock().unwrap();
    if let Some(Value::Object(names)) = guard.properties.get("__col_names") {
        let names_guard = names.lock().unwrap();
        if let ObjectKind::Array(ref elems) = names_guard.kind {
            return elems.iter().map(|value| format!("{}", value)).collect();
        }
    }
    guard
        .properties
        .keys()
        .filter(|key| !key.starts_with("__"))
        .cloned()
        .collect()
}

fn row_item(row: &Arc<Mutex<Object>>, key: Option<&Value>) -> Value {
    let col_names = row_column_names(row);
    let guard = row.lock().unwrap();
    match key {
        Some(Value::F64(index)) => col_names
            .get(*index as usize)
            .and_then(|name| guard.properties.get(name))
            .cloned()
            .unwrap_or(Value::Null),
        Some(Value::I32(index)) => col_names
            .get((*index).max(0) as usize)
            .and_then(|name| guard.properties.get(name))
            .cloned()
            .unwrap_or(Value::Null),
        Some(Value::I64(index)) => col_names
            .get((*index).max(0) as usize)
            .and_then(|name| guard.properties.get(name))
            .cloned()
            .unwrap_or(Value::Null),
        Some(Value::BigInt(index)) => col_names
            .get((*index).max(0) as usize)
            .and_then(|name| guard.properties.get(name))
            .cloned()
            .unwrap_or(Value::Null),
        Some(Value::String(text)) if text.parse::<usize>().is_ok() => col_names
            .get(text.parse::<usize>().unwrap_or(0))
            .and_then(|name| guard.properties.get(name))
            .cloned()
            .unwrap_or(Value::Null),
        Some(other) => {
            let name = format!("{}", other);
            if let Some(value) = guard.properties.get(&name) {
                return value.clone();
            }
            guard
                .properties
                .iter()
                .find(|(key, _)| key.eq_ignore_ascii_case(&name))
                .map(|(_, value)| value.clone())
                .unwrap_or(Value::Null)
        }
        None => Value::Null,
    }
}

pub fn register(vm: &mut VM) {
    // DataTable constructor
    vm.register_host_fn("vybe:data", "dataTableNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let name = args.first().map(|v| format!("{}", v)).unwrap_or_else(|| "Table1".into());
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("DataTable")));
        obj.properties.insert("tablename".into(), Value::String(Arc::from(name.as_str())));
        obj.properties.insert("columns".into(), Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))));
        obj.properties.insert("rows".into(), Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // DataSet constructor
    vm.register_host_fn("vybe:data", "dataSetNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let name = args.first().map(|v| format!("{}", v)).unwrap_or_else(|| "DataSet1".into());
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("DataSet")));
        obj.properties.insert("datasetname".into(), Value::String(Arc::from(name.as_str())));
        obj.properties.insert("tables".into(), Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // DataTable.NewRow()
    vm.register_host_fn("vybe:data", "dataTableNewRow", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("DataRow")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // DataTable.Rows.Add(row)
    vm.register_host_fn("vybe:data", "dataTableAddRow", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(table)) = args.first() {
            let row = args.get(1).cloned().unwrap_or(Value::Null);
            let t = table.lock().unwrap();
            if let Some(Value::Object(rows)) = t.properties.get("rows") {
                let mut r = rows.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = r.kind {
                    elems.push(row);
                }
            }
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:data", "dataTableSelect", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(table)) = args.first() {
            return table.lock().unwrap().properties.get("rows").cloned().unwrap_or_else(|| Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))));
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    vm.register_host_fn("vybe:data", "dataRowItem", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(row)) = args.first() {
            return row_item(row, args.get(1));
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:data", "dataRowIsNull", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(row)) = args.first() {
            return Value::Bool(matches!(row_item(row, args.get(1)), Value::Null));
        }
        Value::Bool(true)
    }));

    vm.register_host_fn("vybe:data", "dataSetTables", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(dataset)) = args.first() {
            return dataset.lock().unwrap().properties.get("tables").cloned().unwrap_or_else(|| Value::Object(Arc::new(Mutex::new(Object::new_array(vec![])))));
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
    }));

    vm.register_host_fn("vybe:data", "dbNullValue", Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Null));
}
