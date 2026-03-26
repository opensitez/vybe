use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::{Object, ObjectKind};
use std::cell::RefCell;

pub fn register(vm: &mut VM) {
    vm.register_host_fn("vybe:fs", "readFile", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        match std::fs::read_to_string(&path) {
            Ok(contents) => Value::String(Rc::from(contents.as_str())),
            Err(e) => Value::String(Rc::from(format!("Error: {}", e).as_str())),
        }
    }));

    vm.register_host_fn("vybe:fs", "readFileBytes", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        match std::fs::read(&path) {
            Ok(bytes) => {
                let vals: Vec<Value> = bytes.iter().map(|b| Value::F64(*b as f64)).collect();
                Value::Object(Rc::new(RefCell::new(Object::new_array(vals))))
            }
            Err(_) => Value::Null,
        }
    }));

    vm.register_host_fn("vybe:fs", "writeFile", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        let data = s(args, 1);
        match std::fs::write(&path, &data) {
            Ok(_) => Value::Bool(true),
            Err(_) => Value::Bool(false),
        }
    }));

    vm.register_host_fn("vybe:fs", "appendFile", Box::new(|args: &[Value]| {
        use std::io::Write;
        let path = s(args, 0);
        let data = s(args, 1);
        match std::fs::OpenOptions::new().create(true).append(true).open(&path) {
            Ok(mut f) => {
                let _ = f.write_all(data.as_bytes());
                Value::Bool(true)
            }
            Err(_) => Value::Bool(false),
        }
    }));

    vm.register_host_fn("vybe:fs", "exists", Box::new(|args: &[Value]| {
        Value::Bool(std::path::Path::new(&s(args, 0)).exists())
    }));

    vm.register_host_fn("vybe:fs", "isFile", Box::new(|args: &[Value]| {
        Value::Bool(std::path::Path::new(&s(args, 0)).is_file())
    }));

    vm.register_host_fn("vybe:fs", "isDir", Box::new(|args: &[Value]| {
        Value::Bool(std::path::Path::new(&s(args, 0)).is_dir())
    }));

    vm.register_host_fn("vybe:fs", "remove", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        let p = std::path::Path::new(&path);
        let ok = if p.is_dir() {
            std::fs::remove_dir_all(p).is_ok()
        } else {
            std::fs::remove_file(p).is_ok()
        };
        Value::Bool(ok)
    }));

    vm.register_host_fn("vybe:fs", "listDir", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        match std::fs::read_dir(&path) {
            Ok(entries) => {
                let items: Vec<Value> = entries
                    .filter_map(|e| e.ok())
                    .map(|e| Value::String(Rc::from(e.file_name().to_string_lossy().as_ref())))
                    .collect();
                Value::Object(Rc::new(RefCell::new(Object::new_array(items))))
            }
            Err(_) => Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))),
        }
    }));

    vm.register_host_fn("vybe:fs", "mkdir", Box::new(|args: &[Value]| {
        Value::Bool(std::fs::create_dir_all(&s(args, 0)).is_ok())
    }));

    vm.register_host_fn("vybe:fs", "fileSize", Box::new(|args: &[Value]| {
        match std::fs::metadata(&s(args, 0)) {
            Ok(m) => Value::F64(m.len() as f64),
            Err(_) => Value::F64(-1.0),
        }
    }));
}

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}
