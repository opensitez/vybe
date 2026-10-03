//! PHP stream metadata not exposed by the generic filesystem surface.
use std::sync::Mutex;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Framework, Value, heap};

static UMASK_LOCK: Mutex<()> = Mutex::new(());

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:filesystem",
        "umask",
        Box::new(|_, args| {
            let _guard = UMASK_LOCK.lock().unwrap();
            #[cfg(unix)]
            {
                let next = args.first().map(|value| value.as_f64() as libc::mode_t);
                // POSIX exposes no read-only umask call. Restore it immediately
                // for PHP's zero-argument form while holding the runtime lock.
                let previous = unsafe { libc::umask(next.unwrap_or(0)) };
                if next.is_none() {
                    unsafe { libc::umask(previous) };
                }
                Value::I64(previous as i64)
            }
            #[cfg(not(unix))]
            {
                let _ = args;
                Value::I64(0)
            }
        }),
    );
    fw.register_host_fn(
        "php:filesystem",
        "chmod",
        Box::new(|_, args| {
            let [Value::String(path), mode, ..] = args else {
                return Value::Bool(false);
            };
            #[cfg(unix)]
            {
                use std::os::unix::fs::PermissionsExt;
                Value::Bool(
                    std::fs::set_permissions(
                        path.as_ref(),
                        std::fs::Permissions::from_mode(mode.as_f64() as u32),
                    )
                    .is_ok(),
                )
            }
            #[cfg(not(unix))]
            {
                let _ = (path, mode);
                Value::Bool(false)
            }
        }),
    );
    fw.register_host_fn(
        "php:filesystem",
        "opendir",
        Box::new(|_, args| {
            let Some(Value::String(path)) = args.first() else {
                return Value::Bool(false);
            };
            let Ok(read_dir) = std::fs::read_dir(path.as_ref()) else {
                return Value::Bool(false);
            };
            let mut names = vec![Value::String(".".into()), Value::String("..".into())];
            names.extend(read_dir.filter_map(|entry| {
                entry.ok().map(|entry| {
                    Value::String(entry.file_name().to_string_lossy().into_owned().into())
                })
            }));
            let mut handle = Object::new();
            handle.kind = ObjectKind::Array(names);
            handle
                .properties
                .insert("__php_directory_index".into(), Value::I64(0));
            Value::Object(heap::alloc(handle))
        }),
    );
    fw.register_host_fn(
        "php:filesystem",
        "readdir",
        Box::new(|_, args| {
            let Some(Value::Object(handle)) = args.first() else {
                return Value::Bool(false);
            };
            let mut handle = handle.lock().unwrap();
            let Some(Value::I64(index)) = handle.properties.get("__php_directory_index") else {
                return Value::Bool(false);
            };
            let index = *index as usize;
            let value = match &handle.kind {
                ObjectKind::Array(names) => names.get(index).cloned(),
                _ => None,
            };
            if value.is_some() {
                handle.properties.insert(
                    "__php_directory_index".into(),
                    Value::I64((index + 1) as i64),
                );
            }
            value.unwrap_or(Value::Bool(false))
        }),
    );
    fw.register_host_fn(
        "php:filesystem",
        "rewinddir",
        Box::new(|_, args| {
            if let Some(Value::Object(handle)) = args.first() {
                let mut handle = handle.lock().unwrap();
                if handle.properties.contains_key("__php_directory_index") {
                    handle
                        .properties
                        .insert("__php_directory_index".into(), Value::I64(0));
                }
            }
            Value::Null
        }),
    );
    fw.register_host_fn(
        "php:filesystem",
        "closedir",
        Box::new(|_, args| {
            if let Some(Value::Object(handle)) = args.first() {
                let mut handle = handle.lock().unwrap();
                handle.properties.remove("__php_directory_index");
                handle.kind = ObjectKind::Array(Vec::new());
            }
            Value::Null
        }),
    );
    fw.register_host_fn("php:filesystem", "fstat", Box::new(|_, args| {
        let Some(Value::Object(stream)) = args.first() else { return Value::Bool(false) };
        let stream = stream.lock().unwrap();
        let Some(Value::String(uri)) = stream.properties.get("__uri") else {
            return Value::Bool(false);
        };
        let sink = stream.properties.get("__sink");
        let is_pipe = matches!(sink, Some(Value::String(value)) if value.as_ref() == "stdout" || value.as_ref() == "stderr");
        let mode = if is_pipe { 0o010000 } else { 0o100600 };
        let size = if uri.starts_with("php://") {
            match stream.properties.get("__buf") {
                Some(Value::String(buffer)) => buffer.len() as i64,
                _ => 0,
            }
        } else {
            std::fs::metadata(uri.as_ref()).map(|meta| meta.len() as i64).unwrap_or(0)
        };
        let names = ["dev", "ino", "mode", "nlink", "uid", "gid", "rdev", "size", "atime", "mtime", "ctime", "blksize", "blocks"];
        let mut result = Object::new();
        result.kind = ObjectKind::Map(Default::default());
        if let ObjectKind::Map(entries) = &mut result.kind {
            for (index, name) in names.iter().enumerate() {
                let value = match *name { "mode" => mode, "size" => size, _ => 0 };
                entries.insert(Value::I64(index as i64), Value::I64(value));
                entries.insert(Value::String((*name).into()), Value::I64(value));
            }
        }
        Value::Object(heap::alloc(result))
    }));
}
