use std::io::{BufRead, Write};
use std::sync::Arc;
use std::sync::Mutex;
use vybe_bytecode::value::Object;
use vybe_bytecode::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "wasi:filesystem",
        "readFile",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            match std::fs::read_to_string(&path) {
                Ok(contents) => Value::String(Arc::from(contents.as_str())),
                Err(e) => Value::String(Arc::from(format!("Error: {}", e).as_str())),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "readFileBytes",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            match std::fs::read(&path) {
                Ok(bytes) => {
                    let vals: Vec<Value> = bytes.iter().map(|b| Value::F64(*b as f64)).collect();
                    Value::Object(vybe_bytecode::heap::alloc(Object::new_array(vals)))
                }
                Err(_) => Value::Null,
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "writeFile",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            let data = s(args, 1);
            match std::fs::write(&path, &data) {
                Ok(_) => Value::Bool(true),
                Err(_) => Value::Bool(false),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "appendFile",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            use std::io::Write;
            let path = s(args, 0);
            let data = s(args, 1);
            match std::fs::OpenOptions::new()
                .create(true)
                .append(true)
                .open(&path)
            {
                Ok(mut f) => {
                    let _ = f.write_all(data.as_bytes());
                    Value::Bool(true)
                }
                Err(_) => Value::Bool(false),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "exists",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::path::Path::new(&s(args, 0)).exists())
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "isFile",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::path::Path::new(&s(args, 0)).is_file())
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "isDir",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::path::Path::new(&s(args, 0)).is_dir())
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "remove",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            let p = std::path::Path::new(&path);
            let ok = if p.is_dir() {
                std::fs::remove_dir_all(p).is_ok()
            } else {
                std::fs::remove_file(p).is_ok()
            };
            Value::Bool(ok)
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "listDir",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            match std::fs::read_dir(&path) {
                Ok(entries) => {
                    let items: Vec<Value> = entries
                        .filter_map(|e| e.ok())
                        .map(|e| Value::String(Arc::from(e.file_name().to_string_lossy().as_ref())))
                        .collect();
                    Value::Object(vybe_bytecode::heap::alloc(Object::new_array(items)))
                }
                Err(_) => Value::Object(vybe_bytecode::heap::alloc(Object::new_array(vec![]))),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "mkdir",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::fs::create_dir_all(&s(args, 0)).is_ok())
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "fileSize",
        Box::new(
            |_ctx: &mut HostContext, args: &[Value]| match std::fs::metadata(&s(args, 0)) {
                Ok(m) => Value::F64(m.len() as f64),
                Err(_) => Value::F64(-1.0),
            },
        ),
    );

    // rename(oldPath, newPath)
    vm.register_host_fn(
        "wasi:filesystem",
        "rename",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let old = s(args, 0);
            let new = s(args, 1);
            Value::Bool(std::fs::rename(&old, &new).is_ok())
        }),
    );

    // copy(src, dest)
    vm.register_host_fn(
        "wasi:filesystem",
        "copy",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let src = s(args, 0);
            let dest = s(args, 1);
            Value::Bool(std::fs::copy(&src, &dest).is_ok())
        }),
    );

    // stat(path) → object { size, isFile, isDir, modified }
    vm.register_host_fn(
        "wasi:filesystem",
        "stat",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            match std::fs::metadata(&path) {
                Ok(m) => {
                    let mut obj = Object::new();
                    obj.properties
                        .insert("size".into(), Value::F64(m.len() as f64));
                    obj.properties
                        .insert("isFile".into(), Value::Bool(m.is_file()));
                    obj.properties
                        .insert("isDir".into(), Value::Bool(m.is_dir()));
                    if let Ok(modified) = m.modified() {
                        let ms = modified
                            .duration_since(std::time::UNIX_EPOCH)
                            .unwrap_or_default()
                            .as_millis();
                        obj.properties
                            .insert("modified".into(), Value::F64(ms as f64));
                    }
                    Value::Object(vybe_bytecode::heap::alloc(obj))
                }
                Err(_) => Value::Null,
            }
        }),
    );

    // readDir(path) → array of { name, isFile, isDir }
    vm.register_host_fn(
        "wasi:filesystem",
        "readDirEntries",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            match std::fs::read_dir(&path) {
                Ok(entries) => {
                    let items: Vec<Value> = entries
                        .filter_map(|e| e.ok())
                        .map(|e| {
                            let mut obj = Object::new();
                            obj.properties.insert(
                                "name".into(),
                                Value::String(Arc::from(e.file_name().to_string_lossy().as_ref())),
                            );
                            if let Ok(ft) = e.file_type() {
                                obj.properties
                                    .insert("isFile".into(), Value::Bool(ft.is_file()));
                                obj.properties
                                    .insert("isDir".into(), Value::Bool(ft.is_dir()));
                            }
                            Value::Object(vybe_bytecode::heap::alloc(obj))
                        })
                        .collect();
                    Value::Object(vybe_bytecode::heap::alloc(Object::new_array(items)))
                }
                Err(_) => Value::Object(vybe_bytecode::heap::alloc(Object::new_array(vec![]))),
            }
        }),
    );

    // --- Path functions ---

    vm.register_host_fn(
        "wasi:filesystem",
        "pathCombine",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let mut path = std::path::PathBuf::from(s(args, 0));
            for i in 1..args.len() {
                path.push(s(args, i));
            }
            Value::String(Arc::from(path.to_string_lossy().as_ref()))
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "pathGetFileName",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let input = s(args, 0);
            let p = std::path::Path::new(&input);
            Value::String(Arc::from(
                p.file_name().unwrap_or_default().to_string_lossy().as_ref(),
            ))
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "pathGetExtension",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let input = s(args, 0);
            let p = std::path::Path::new(&input);
            Value::String(Arc::from(
                p.extension().unwrap_or_default().to_string_lossy().as_ref(),
            ))
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "pathGetDirectory",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let input = s(args, 0);
            let p = std::path::Path::new(&input);
            Value::String(Arc::from(
                p.parent()
                    .map(|p| p.to_string_lossy().into_owned())
                    .unwrap_or_default()
                    .as_str(),
            ))
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "pathGetFileNameWithoutExt",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let input = s(args, 0);
            let p = std::path::Path::new(&input);
            Value::String(Arc::from(
                p.file_stem().unwrap_or_default().to_string_lossy().as_ref(),
            ))
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "pathChangeExtension",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let mut p = std::path::PathBuf::from(s(args, 0));
            p.set_extension(s(args, 1).trim_start_matches('.'));
            Value::String(Arc::from(p.to_string_lossy().as_ref()))
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "pathGetFullPath",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match std::fs::canonicalize(s(args, 0)) {
                Ok(p) => Value::String(Arc::from(p.to_string_lossy().as_ref())),
                Err(_) => Value::String(Arc::from(s(args, 0).as_str())),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "pathGetTempPath",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            Value::String(Arc::from(std::env::temp_dir().to_string_lossy().as_ref()))
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "pathHasExtension",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let p = s(args, 0);
            Value::Bool(std::path::Path::new(p.as_str()).extension().is_some())
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem",
        "pathIsRooted",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::path::Path::new(s(args, 0).as_str()).is_absolute())
        }),
    );

    // -- VB6 file handle I/O --

    // openFile(path, mode, fileNumber) → null
    vm.register_host_fn(
        "wasi:filesystem",
        "openFile",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            let mode = s(args, 1);
            let fnum = args.get(2).map(|v| v.as_f64() as i32).unwrap_or(1);
            let handle = match mode.as_str() {
                "Input" => match std::fs::File::open(&path) {
                    Ok(f) => FileHandle {
                        reader: Some(std::io::BufReader::new(f)),
                        writer: None,
                    },
                    Err(_) => return Value::Null,
                },
                "Output" => match std::fs::File::create(&path) {
                    Ok(f) => FileHandle {
                        reader: None,
                        writer: Some(std::io::BufWriter::new(f)),
                    },
                    Err(_) => return Value::Null,
                },
                "Append" => {
                    match std::fs::OpenOptions::new()
                        .append(true)
                        .create(true)
                        .open(&path)
                    {
                        Ok(f) => FileHandle {
                            reader: None,
                            writer: Some(std::io::BufWriter::new(f)),
                        },
                        Err(_) => return Value::Null,
                    }
                }
                _ => return Value::Null,
            };
            if let Ok(mut handles) = FILE_HANDLES.lock() {
                handles.insert(fnum, handle);
            }
            Value::Null
        }),
    );

    // closeFile(fileNumber) → null (-1 = close all)
    vm.register_host_fn(
        "wasi:filesystem",
        "closeFile",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let fnum = args.first().map(|v| v.as_i32()).unwrap_or(-1);
            if let Ok(mut handles) = FILE_HANDLES.lock() {
                if fnum == -1 {
                    handles.clear();
                } else {
                    handles.remove(&fnum);
                }
            }
            Value::Null
        }),
    );

    // printFile(fileNumber, items...) → null
    vm.register_host_fn(
        "wasi:filesystem",
        "printFile",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let fnum = args.first().map(|v| v.as_i32()).unwrap_or(1);
            let text: String = args[1..]
                .iter()
                .map(|v| format!("{}", v))
                .collect::<Vec<_>>()
                .join("");
            if let Ok(mut handles) = FILE_HANDLES.lock() {
                if let Some(h) = handles.get_mut(&fnum) {
                    if let Some(ref mut w) = h.writer {
                        let _ = writeln!(w, "{}", text);
                        let _ = w.flush();
                    }
                }
            }
            Value::Null
        }),
    );

    // writeFile(fileNumber, items...) → null (CSV-style with quotes)
    vm.register_host_fn(
        "wasi:filesystem",
        "writeFile_handle",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let fnum = args.first().map(|v| v.as_i32()).unwrap_or(1);
            let parts: Vec<String> = args[1..]
                .iter()
                .map(|v| {
                    let s = format!("{}", v);
                    if s.contains(',') || s.contains('"') {
                        format!("\"{}\"", s.replace('"', "\"\""))
                    } else {
                        s
                    }
                })
                .collect();
            if let Ok(mut handles) = FILE_HANDLES.lock() {
                if let Some(h) = handles.get_mut(&fnum) {
                    if let Some(ref mut w) = h.writer {
                        let _ = writeln!(w, "{}", parts.join(","));
                        let _ = w.flush();
                    }
                }
            }
            Value::Null
        }),
    );

    // lineInput(fileNumber) → string (one line)
    vm.register_host_fn(
        "wasi:filesystem",
        "lineInput",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let fnum = args.first().map(|v| v.as_i32()).unwrap_or(1);
            if let Ok(mut handles) = FILE_HANDLES.lock() {
                if let Some(h) = handles.get_mut(&fnum) {
                    if let Some(ref mut r) = h.reader {
                        let mut line = String::new();
                        if r.read_line(&mut line).is_ok() {
                            return Value::String(Arc::from(
                                line.trim_end_matches('\n').trim_end_matches('\r'),
                            ));
                        }
                    }
                }
            }
            Value::String(Arc::from(""))
        }),
    );

    // inputFile(fileNumber) → array of comma-separated values from one line
    vm.register_host_fn(
        "wasi:filesystem",
        "inputFile",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let fnum = args.first().map(|v| v.as_i32()).unwrap_or(1);
            if let Ok(mut handles) = FILE_HANDLES.lock() {
                if let Some(h) = handles.get_mut(&fnum) {
                    if let Some(ref mut r) = h.reader {
                        let mut line = String::new();
                        if r.read_line(&mut line).is_ok() {
                            let vals: Vec<Value> = line
                                .trim()
                                .split(',')
                                .map(|s| Value::String(Arc::from(s.trim())))
                                .collect();
                            return Value::Object(vybe_bytecode::heap::alloc(Object::new_array(vals)));
                        }
                    }
                }
            }
            Value::Object(vybe_bytecode::heap::alloc(Object::new_array(vec![])))
        }),
    );
}

struct FileHandle {
    reader: Option<std::io::BufReader<std::fs::File>>,
    writer: Option<std::io::BufWriter<std::fs::File>>,
}

static FILE_HANDLES: std::sync::LazyLock<Mutex<std::collections::HashMap<i32, FileHandle>>> =
    std::sync::LazyLock::new(|| Mutex::new(std::collections::HashMap::new()));

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}
