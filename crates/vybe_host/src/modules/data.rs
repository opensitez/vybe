//! System.Data — DataTable, DataSet (simplified objects)

use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // DataTable constructor
    vm.register_host_fn("vybe:data", "dataTableNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let name = args.first().map(|v| format!("{}", v)).unwrap_or_else(|| "Table1".into());
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("DataTable")));
        obj.properties.insert("tablename".into(), Value::String(Rc::from(name.as_str())));
        obj.properties.insert("columns".into(), Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))));
        obj.properties.insert("rows".into(), Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // DataSet constructor
    vm.register_host_fn("vybe:data", "dataSetNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let name = args.first().map(|v| format!("{}", v)).unwrap_or_else(|| "DataSet1".into());
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("DataSet")));
        obj.properties.insert("datasetname".into(), Value::String(Rc::from(name.as_str())));
        obj.properties.insert("tables".into(), Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // DataTable.NewRow()
    vm.register_host_fn("vybe:data", "dataTableNewRow", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("DataRow")));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // DataTable.Rows.Add(row)
    vm.register_host_fn("vybe:data", "dataTableAddRow", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(table)) = args.first() {
            let row = args.get(1).cloned().unwrap_or(Value::Null);
            let t = table.borrow();
            if let Some(Value::Object(rows)) = t.properties.get("rows") {
                let mut r = rows.borrow_mut();
                if let ObjectKind::Array(ref mut elems) = r.kind {
                    elems.push(row);
                }
            }
        }
        Value::Null
    }));
}
