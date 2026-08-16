use std::io::{BufRead, Write};
use std::sync::Arc;
use std::sync::Mutex;
use vybe_runtime::value::Object;
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, VM, ValType, Value};

/// One path argument — the shape of most of this module.
fn path() -> Vec<ValType> {
    vec![ValType::String]
}

/// Register a `wasi:filesystem` function WITH its signature.
///
/// No resource binding, unlike `io.rs`: this surface is path-based, and a path
/// is a string rather than a handle the host owns. `openFile`/`closeFile` are
/// the exception — they pass a file HANDLE — but it travels as a plain value
/// here, so there is no resource type to bind to and claiming one would be a
/// model this module does not have.
fn fs_fn(
    vm: &mut VM,
    name: &str,
    kebab: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("wasi:filesystem", name, call).with_sig(FuncSig {
            name: kebab.to_string(),
            params,
            results,
        }),
    );
}

/// `pathCombine`, `printFile`, `writeFile_handle`, `lineInput` and `inputFile`
/// are deliberately left UNDECLARED.
///
/// The first three are variadic — `pathCombine` is `Path.Combine(a, b, c…)` and
/// takes however many segments it is given. `lineInput` is emitted with argc 1,
/// 2 AND 3 by the Pascal frontend (`ReadLn` with a varying number of targets),
/// which is the same thing seen from the call side. `inputFile`'s closure reads
/// no arguments and has no measured call site, so its arity is genuinely
/// unknown — and UNKNOWN is what leaving it undeclared says.
pub fn register(vm: &mut VM) {
    fs_fn(
        vm,
        "readFile",
        "read-file",
        path(),
        // The closure answers `"Error: …"` as a STRING on failure rather than a
        // `result`, so `String` is the honest declaration of what it returns.
        // Declaring `result<string, _>` would describe a shape no caller can
        // discriminate — that is a real gap, and naming it here is better than
        // a signature that quietly claims otherwise.
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            match std::fs::read_to_string(&path) {
                Ok(contents) => Value::String(Arc::from(contents.as_str())),
                Err(e) => Value::String(Arc::from(format!("Error: {}", e).as_str())),
            }
        }),
    );

    fs_fn(
        vm,
        "readFileBytes",
        "read-file-bytes",
        path(),
        // `null` on failure, so the answer is genuinely optional here — unlike
        // `readFile` next door, which folds the error into the string.
        vec![ValType::Option(Box::new(ValType::List(Box::new(
            ValType::I32,
        ))))],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            match std::fs::read(&path) {
                Ok(bytes) => {
                    let vals: Vec<Value> = bytes.iter().map(|b| Value::F64(*b as f64)).collect();
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(vals)))
                }
                Err(_) => Value::Null,
            }
        }),
    );

    fs_fn(
        vm,
        "writeFile",
        "write-file",
        vec![ValType::String, ValType::String],
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            let data = s(args, 1);
            match std::fs::write(&path, &data) {
                Ok(_) => Value::Bool(true),
                Err(_) => Value::Bool(false),
            }
        }),
    );

    fs_fn(
        vm,
        "appendFile",
        "append-file",
        vec![ValType::String, ValType::String],
        vec![ValType::Bool],
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

    fs_fn(
        vm,
        "exists",
        "exists",
        path(),
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::path::Path::new(&s(args, 0)).exists())
        }),
    );

    fs_fn(
        vm,
        "isFile",
        "is-file",
        path(),
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::path::Path::new(&s(args, 0)).is_file())
        }),
    );

    fs_fn(
        vm,
        "isDir",
        "is-dir",
        path(),
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::path::Path::new(&s(args, 0)).is_dir())
        }),
    );

    fs_fn(
        vm,
        "remove",
        "remove",
        path(),
        vec![ValType::Bool],
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

    fs_fn(
        vm,
        "listDir",
        "list-dir",
        path(),
        vec![ValType::List(Box::new(ValType::String))],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let path = s(args, 0);
            match std::fs::read_dir(&path) {
                Ok(entries) => {
                    let items: Vec<Value> = entries
                        .filter_map(|e| e.ok())
                        .map(|e| Value::String(Arc::from(e.file_name().to_string_lossy().as_ref())))
                        .collect();
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
                }
                Err(_) => Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![]))),
            }
        }),
    );

    fs_fn(
        vm,
        "mkdir",
        "mkdir",
        path(),
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::fs::create_dir_all(&s(args, 0)).is_ok())
        }),
    );

    fs_fn(
        vm,
        "fileSize",
        "file-size",
        path(),
        vec![ValType::F64],
        Box::new(
            |_ctx: &mut HostContext, args: &[Value]| match std::fs::metadata(&s(args, 0)) {
                Ok(m) => Value::F64(m.len() as f64),
                Err(_) => Value::F64(-1.0),
            },
        ),
    );

    // rename(oldPath, newPath)
    fs_fn(
        vm,
        "rename",
        "rename",
        vec![ValType::String, ValType::String],
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let old = s(args, 0);
            let new = s(args, 1);
            Value::Bool(std::fs::rename(&old, &new).is_ok())
        }),
    );

    // copy(src, dest)
    fs_fn(
        vm,
        "copy",
        "copy",
        vec![ValType::String, ValType::String],
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let src = s(args, 0);
            let dest = s(args, 1);
            Value::Bool(std::fs::copy(&src, &dest).is_ok())
        }),
    );

    // stat(path) → object { size, isFile, isDir, modified }
    fs_fn(
        vm,
        "stat",
        "stat",
        path(),
        // A record of fields, not a scalar — `Any` because `ValType::Record`
        // would have to name and order every field, and this answers an object
        // whose shape belongs to the caller's language, not to WIT.
        vec![ValType::Any],
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
                    Value::Object(vybe_runtime::heap::alloc(obj))
                }
                Err(_) => Value::Null,
            }
        }),
    );

    // readDir(path) → array of { name, isFile, isDir }
    fs_fn(
        vm,
        "readDirEntries",
        "read-dir-entries",
        path(),
        vec![ValType::List(Box::new(ValType::Any))],
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
                            Value::Object(vybe_runtime::heap::alloc(obj))
                        })
                        .collect();
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
                }
                Err(_) => Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![]))),
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

    fs_fn(
        vm,
        "pathGetFileName",
        "path-get-file-name",
        path(),
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let input = s(args, 0);
            let p = std::path::Path::new(&input);
            Value::String(Arc::from(
                p.file_name().unwrap_or_default().to_string_lossy().as_ref(),
            ))
        }),
    );

    fs_fn(
        vm,
        "pathGetExtension",
        "path-get-extension",
        path(),
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let input = s(args, 0);
            let p = std::path::Path::new(&input);
            Value::String(Arc::from(
                p.extension().unwrap_or_default().to_string_lossy().as_ref(),
            ))
        }),
    );

    fs_fn(
        vm,
        "pathGetDirectory",
        "path-get-directory",
        path(),
        vec![ValType::String],
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

    fs_fn(
        vm,
        "pathGetFileNameWithoutExt",
        "path-get-file-name-without-ext",
        path(),
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let input = s(args, 0);
            let p = std::path::Path::new(&input);
            Value::String(Arc::from(
                p.file_stem().unwrap_or_default().to_string_lossy().as_ref(),
            ))
        }),
    );

    fs_fn(
        vm,
        "pathChangeExtension",
        "path-change-extension",
        vec![ValType::String, ValType::String],
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let mut p = std::path::PathBuf::from(s(args, 0));
            p.set_extension(s(args, 1).trim_start_matches('.'));
            Value::String(Arc::from(p.to_string_lossy().as_ref()))
        }),
    );

    fs_fn(
        vm,
        "pathGetFullPath",
        "path-get-full-path",
        path(),
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            match std::fs::canonicalize(s(args, 0)) {
                Ok(p) => Value::String(Arc::from(p.to_string_lossy().as_ref())),
                Err(_) => Value::String(Arc::from(s(args, 0).as_str())),
            }
        }),
    );

    fs_fn(
        vm,
        "pathGetTempPath",
        "path-get-temp-path",
        // The only zero-parameter function here, and the call sites agree.
        vec![],
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            Value::String(Arc::from(std::env::temp_dir().to_string_lossy().as_ref()))
        }),
    );

    fs_fn(
        vm,
        "pathHasExtension",
        "path-has-extension",
        path(),
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let p = s(args, 0);
            Value::Bool(std::path::Path::new(p.as_str()).extension().is_some())
        }),
    );

    fs_fn(
        vm,
        "pathIsRooted",
        "path-is-rooted",
        path(),
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(std::path::Path::new(s(args, 0).as_str()).is_absolute())
        }),
    );

    // -- VB6 file handle I/O --

    // openFile(path, mode, fileNumber) → null
    fs_fn(
        vm,
        "openFile",
        "open-file",
        // (path, mode, fileNumber) — three, which is what every LIVE call site
        // passes (`builtins.rs`, `statements.rs`). One 2-arg site exists,
        // `primitives::io::emit_open_file`, but that helper has zero callers —
        // it and its four `emit_*_file` neighbours are definitions only — so no
        // reachable route disagrees.
        vec![ValType::String, ValType::String, ValType::String],
        vec![ValType::F64],
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
    fs_fn(
        vm,
        "closeFile",
        "close-file",
        // ONE parameter, from the CALL SITES — not from the closure, which
        // reads none. Same shape as `wasi:io`'s `flush` and `pollable.ready`:
        // the host keeps one table and does not need the handle to find the
        // entry, so the closure never consults what every caller passes. The
        // closure is a lower bound on the contract; the callers are the
        // contract.
        vec![ValType::F64],
        vec![ValType::Bool],
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
                            return Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                                vals,
                            )));
                        }
                    }
                }
            }
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![])))
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
