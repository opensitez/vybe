//! Resolve the PHP request store without expanding its lazy initialization at
//! every variable access. The caller supplies the existing globalThis object;
//! no process-local cache can outlive a VM reset or detach included modules.
use vybe_runtime::value::Object;
use vybe_runtime::{Framework, Value, heap};

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:globals",
        "store",
        Box::new(|_, args| {
            let Some(Value::Object(global_this)) = args.first() else {
                return Value::Null;
            };
            let mut global_this = global_this.lock().unwrap();
            const KEY: &str = "__vybe_php_globals";
            if let Some(store) = global_this.properties.get(KEY) {
                if !matches!(store, Value::Null | Value::Undefined) {
                    return store.clone();
                }
            }
            let store = Value::Object(heap::alloc(Object::new()));
            global_this.properties.insert(KEY.to_owned(), store.clone());
            store
        }),
    );
}
