use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use vybe_bytecode::value::{Object, ObjectKind};
use std::cell::RefCell;

pub fn register(vm: &mut VM) {
    vm.register_host_fn("wasi:filesystem", "readFile", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        match std::fs::read_to_string(&path) {
            Ok(contents) => Value::String(Rc::from(contents.as_str())),
            Err(e) => Value::String(Rc::from(format!("Error: {}", e).as_str())),
        }
    }));

    vm.register_host_fn("wasi:filesystem", "readFileBytes", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        match std::fs::read(&path) {
            Ok(bytes) => {
                let vals: Vec<Value> = bytes.iter().map(|b| Value::F64(*b as f64)).collect();
                Value::Object(Rc::new(RefCell::new(Object::new_array(vals))))
            }
            Err(_) => Value::Null,
        }
    }));

    vm.register_host_fn("wasi:filesystem", "writeFile", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        let data = s(args, 1);
        match std::fs::write(&path, &data) {
            Ok(_) => Value::Bool(true),
            Err(_) => Value::Bool(false),
        }
    }));

    vm.register_host_fn("wasi:filesystem", "appendFile", Box::new(|args: &[Value]| {
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

    vm.register_host_fn("wasi:filesystem", "exists", Box::new(|args: &[Value]| {
        Value::Bool(std::path::Path::new(&s(args, 0)).exists())
    }));

    vm.register_host_fn("wasi:filesystem", "isFile", Box::new(|args: &[Value]| {
        Value::Bool(std::path::Path::new(&s(args, 0)).is_file())
    }));

    vm.register_host_fn("wasi:filesystem", "isDir", Box::new(|args: &[Value]| {
        Value::Bool(std::path::Path::new(&s(args, 0)).is_dir())
    }));

    vm.register_host_fn("wasi:filesystem", "remove", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        let p = std::path::Path::new(&path);
        let ok = if p.is_dir() {
            std::fs::remove_dir_all(p).is_ok()
        } else {
            std::fs::remove_file(p).is_ok()
        };
        Value::Bool(ok)
    }));

    vm.register_host_fn("wasi:filesystem", "listDir", Box::new(|args: &[Value]| {
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

    vm.register_host_fn("wasi:filesystem", "mkdir", Box::new(|args: &[Value]| {
        Value::Bool(std::fs::create_dir_all(&s(args, 0)).is_ok())
    }));

    vm.register_host_fn("wasi:filesystem", "fileSize", Box::new(|args: &[Value]| {
        match std::fs::metadata(&s(args, 0)) {
            Ok(m) => Value::F64(m.len() as f64),
            Err(_) => Value::F64(-1.0),
        }
    }));

    // rename(oldPath, newPath)
    vm.register_host_fn("wasi:filesystem", "rename", Box::new(|args: &[Value]| {
        let old = s(args, 0);
        let new = s(args, 1);
        Value::Bool(std::fs::rename(&old, &new).is_ok())
    }));

    // copy(src, dest)
    vm.register_host_fn("wasi:filesystem", "copy", Box::new(|args: &[Value]| {
        let src = s(args, 0);
        let dest = s(args, 1);
        Value::Bool(std::fs::copy(&src, &dest).is_ok())
    }));

    // stat(path) → object { size, isFile, isDir, modified }
    vm.register_host_fn("wasi:filesystem", "stat", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        match std::fs::metadata(&path) {
            Ok(m) => {
                let mut obj = Object::new();
                obj.properties.insert("size".into(), Value::F64(m.len() as f64));
                obj.properties.insert("isFile".into(), Value::Bool(m.is_file()));
                obj.properties.insert("isDir".into(), Value::Bool(m.is_dir()));
                if let Ok(modified) = m.modified() {
                    let ms = modified.duration_since(std::time::UNIX_EPOCH).unwrap_or_default().as_millis();
                    obj.properties.insert("modified".into(), Value::F64(ms as f64));
                }
                Value::Object(Rc::new(RefCell::new(obj)))
            }
            Err(_) => Value::Null,
        }
    }));

    // readDir(path) → array of { name, isFile, isDir }
    vm.register_host_fn("wasi:filesystem", "readDirEntries", Box::new(|args: &[Value]| {
        let path = s(args, 0);
        match std::fs::read_dir(&path) {
            Ok(entries) => {
                let items: Vec<Value> = entries.filter_map(|e| e.ok()).map(|e| {
                    let mut obj = Object::new();
                    obj.properties.insert("name".into(), Value::String(Rc::from(e.file_name().to_string_lossy().as_ref())));
                    if let Ok(ft) = e.file_type() {
                        obj.properties.insert("isFile".into(), Value::Bool(ft.is_file()));
                        obj.properties.insert("isDir".into(), Value::Bool(ft.is_dir()));
                    }
                    Value::Object(Rc::new(RefCell::new(obj)))
                }).collect();
                Value::Object(Rc::new(RefCell::new(Object::new_array(items))))
            }
            Err(_) => Value::Object(Rc::new(RefCell::new(Object::new_array(vec![])))),
        }
    }));

    // --- Path functions ---

    vm.register_host_fn("wasi:filesystem", "pathCombine", Box::new(|args: &[Value]| {
        let mut path = std::path::PathBuf::from(s(args, 0));
        for i in 1..args.len() { path.push(s(args, i)); }
        Value::String(Rc::from(path.to_string_lossy().as_ref()))
    }));

    vm.register_host_fn("wasi:filesystem", "pathGetFileName", Box::new(|args: &[Value]| {
        let input = s(args, 0);
        let p = std::path::Path::new(&input);
        Value::String(Rc::from(p.file_name().unwrap_or_default().to_string_lossy().as_ref()))
    }));

    vm.register_host_fn("wasi:filesystem", "pathGetExtension", Box::new(|args: &[Value]| {
        let input = s(args, 0);
        let p = std::path::Path::new(&input);
        Value::String(Rc::from(p.extension().unwrap_or_default().to_string_lossy().as_ref()))
    }));

    vm.register_host_fn("wasi:filesystem", "pathGetDirectory", Box::new(|args: &[Value]| {
        let input = s(args, 0);
        let p = std::path::Path::new(&input);
        Value::String(Rc::from(p.parent().map(|p| p.to_string_lossy().into_owned()).unwrap_or_default().as_str()))
    }));

    vm.register_host_fn("wasi:filesystem", "pathGetFileNameWithoutExt", Box::new(|args: &[Value]| {
        let input = s(args, 0);
        let p = std::path::Path::new(&input);
        Value::String(Rc::from(p.file_stem().unwrap_or_default().to_string_lossy().as_ref()))
    }));

    vm.register_host_fn("wasi:filesystem", "pathChangeExtension", Box::new(|args: &[Value]| {
        let mut p = std::path::PathBuf::from(s(args, 0));
        p.set_extension(s(args, 1).trim_start_matches('.'));
        Value::String(Rc::from(p.to_string_lossy().as_ref()))
    }));

    vm.register_host_fn("wasi:filesystem", "pathGetFullPath", Box::new(|args: &[Value]| {
        match std::fs::canonicalize(s(args, 0)) {
            Ok(p) => Value::String(Rc::from(p.to_string_lossy().as_ref())),
            Err(_) => Value::String(Rc::from(s(args, 0).as_str())),
        }
    }));

    vm.register_host_fn("wasi:filesystem", "pathGetTempPath", Box::new(|_args: &[Value]| {
        Value::String(Rc::from(std::env::temp_dir().to_string_lossy().as_ref()))
    }));
}

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}
