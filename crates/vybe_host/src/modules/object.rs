use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // Object.keys(obj) → array of property name strings
    vm.register_host_fn("vybe:object", "keys", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let keys: Vec<Value> = o.properties.keys()
                .filter(|k| *k != "length") // exclude internal 'length' for arrays
                .map(|k| Value::String(Rc::from(k.as_str())))
                .collect();
            return Value::Object(Rc::new(RefCell::new(Object::new_array(keys))));
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));

    // Object.values(obj) → array of property values
    vm.register_host_fn("vybe:object", "values", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let vals: Vec<Value> = o.properties.values().cloned().collect();
            return Value::Object(Rc::new(RefCell::new(Object::new_array(vals))));
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));

    // Object.entries(obj) → array of [key, value] pairs
    vm.register_host_fn("vybe:object", "entries", Box::new(|args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.borrow();
            let entries: Vec<Value> = o.properties.iter()
                .map(|(k, v)| {
                    Value::Object(Rc::new(RefCell::new(Object::new_array(vec![
                        Value::String(Rc::from(k.as_str())),
                        v.clone(),
                    ]))))
                })
                .collect();
            return Value::Object(Rc::new(RefCell::new(Object::new_array(entries))));
        }
        Value::Object(Rc::new(RefCell::new(Object::new_array(vec![]))))
    }));

    // Object.assign(target, source) → target with source props copied
    vm.register_host_fn("vybe:object", "assign", Box::new(|args: &[Value]| {
        if let (Some(Value::Object(target)), Some(Value::Object(source))) = (args.first(), args.get(1)) {
            let src = source.borrow();
            let mut tgt = target.borrow_mut();
            for (k, v) in &src.properties {
                tgt.properties.insert(k.clone(), v.clone());
            }
        }
        args.first().cloned().unwrap_or(Value::Null)
    }));
}
