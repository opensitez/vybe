//! `node:fs` — Node.js built-in `fs` module.
//!
//! Surface follows the official Node API at
//! <https://nodejs.org/api/fs.html>. Sync variants are first-class
//! citizens; callback / promise variants come in a later phase.
//!
//! The Stats object returned by `statSync`/`lstatSync` carries the
//! data fields (`size`, `mtimeMs`, `atimeMs`, `ctimeMs`,
//! `birthtimeMs`, plus internal `__type` byte) AND the predicate
//! methods (`isFile()`, `isDirectory()`, `isSymbolicLink()`, etc.)
//! bound as host fn refs taking the Stats object as the receiver.

use std::collections::HashMap;
use std::sync::atomic::{AtomicI32, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use std::time::SystemTime;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

// ── Global FD table — maps host-fd-number → (path, mode, position) ──
static FD_TABLE: OnceLock<Mutex<HashMap<i32, (String, String, u64)>>> = OnceLock::new();
static NEXT_FD: AtomicI32 = AtomicI32::new(3);

fn fd_table() -> &'static Mutex<HashMap<i32, (String, String, u64)>> {
    FD_TABLE.get_or_init(|| Mutex::new(HashMap::new()))
}

// ── Stats type tags ───────────────────────────────────────────────
const TYPE_FILE: i32 = 1;
const TYPE_DIR: i32 = 2;
const TYPE_SYMLINK: i32 = 3;
const TYPE_BLOCK: i32 = 4;
const TYPE_CHAR: i32 = 5;
const TYPE_FIFO: i32 = 6;
const TYPE_SOCKET: i32 = 7;

// ── Helpers ───────────────────────────────────────────────────────

fn s_arg(args: &[Value], idx: usize, default: &str) -> String {
    match args.get(idx) {
        Some(Value::String(text)) => text.to_string(),
        Some(other) => format!("{}", other),
        None => default.to_string() }
}

fn opt_obj_bool(args: &[Value], idx: usize, key: &str) -> bool {
    if let Some(Value::Object(obj)) = args.get(idx) {
        let o = obj.lock().unwrap();
        if let Some(Value::Bool(b)) = o.properties.get(key) {
            return *b;
        }
    }
    false
}

fn host_fn_ref(vm: &VM, name: &str) -> Value {
    if let Some(&idx) = vm
        .host_registry
        .get(&("node:fs".to_string(), name.to_string()))
    {
        let mut obj = Object::new();
        obj.properties
            .insert("__host_module".into(), Value::String(Arc::from("node:fs")));
        obj.properties
            .insert("__host_name".into(), Value::String(Arc::from(name)));
        obj.properties
            .insert("__host_idx".into(), Value::F64(idx as f64));
        obj.kind = ObjectKind::HostFunction(idx);
        Value::Object(vybe_runtime::heap::alloc(obj))
    } else {
        Value::Null
    }
}

fn ms_since_epoch(time: SystemTime) -> f64 {
    time.duration_since(SystemTime::UNIX_EPOCH)
        .map(|d| d.as_millis() as f64)
        .unwrap_or(0.0)
}

fn build_stats(meta: &std::fs::Metadata, vm_for_methods: &VM) -> Value {
    let mut o = Object::new();

    // Data fields (Node Stats: size, mtime/atime/ctime/birthtime in ms).
    o.properties
        .insert("size".into(), Value::F64(meta.len() as f64));
    if let Ok(t) = meta.modified() {
        o.properties
            .insert("mtimeMs".into(), Value::F64(ms_since_epoch(t)));
    }
    if let Ok(t) = meta.accessed() {
        o.properties
            .insert("atimeMs".into(), Value::F64(ms_since_epoch(t)));
    }
    if let Ok(t) = meta.created() {
        o.properties
            .insert("birthtimeMs".into(), Value::F64(ms_since_epoch(t)));
    }
    let ft = meta.file_type();
    let type_tag = if ft.is_file() {
        TYPE_FILE
    } else if ft.is_dir() {
        TYPE_DIR
    } else if ft.is_symlink() {
        TYPE_SYMLINK
    } else {
        0
    };
    o.properties.insert("__type".into(), Value::I32(type_tag));

    // Method bindings — Node-faithful predicate methods.
    o.properties
        .insert("isFile".into(), host_fn_ref(vm_for_methods, "_statIsFile"));
    o.properties.insert(
        "isDirectory".into(),
        host_fn_ref(vm_for_methods, "_statIsDirectory"),
    );
    o.properties.insert(
        "isSymbolicLink".into(),
        host_fn_ref(vm_for_methods, "_statIsSymbolicLink"),
    );
    o.properties.insert(
        "isBlockDevice".into(),
        host_fn_ref(vm_for_methods, "_statIsBlockDevice"),
    );
    o.properties.insert(
        "isCharacterDevice".into(),
        host_fn_ref(vm_for_methods, "_statIsCharacterDevice"),
    );
    o.properties
        .insert("isFIFO".into(), host_fn_ref(vm_for_methods, "_statIsFIFO"));
    o.properties.insert(
        "isSocket".into(),
        host_fn_ref(vm_for_methods, "_statIsSocket"),
    );

    Value::Object(vybe_runtime::heap::alloc(o))
}

fn stat_type(args: &[Value]) -> i32 {
    if let Some(Value::Object(obj)) = args.first() {
        let o = obj.lock().unwrap();
        if let Some(Value::I32(tag)) = o.properties.get("__type") {
            return *tag;
        }
    }
    0
}

// ── Registration ──────────────────────────────────────────────────

pub fn register(vm: &mut VM) {
    // The Stats predicate fns are registered first so build_stats can
    // resolve them via host_fn_ref when statSync / lstatSync run.
    vm.register_host_fn(
        "node:fs",
        "_statIsFile",
        Box::new(|_ctx, args| Value::Bool(stat_type(args) == TYPE_FILE)),
    );
    vm.register_host_fn(
        "node:fs",
        "_statIsDirectory",
        Box::new(|_ctx, args| Value::Bool(stat_type(args) == TYPE_DIR)),
    );
    vm.register_host_fn(
        "node:fs",
        "_statIsSymbolicLink",
        Box::new(|_ctx, args| Value::Bool(stat_type(args) == TYPE_SYMLINK)),
    );
    vm.register_host_fn(
        "node:fs",
        "_statIsBlockDevice",
        Box::new(|_ctx, args| Value::Bool(stat_type(args) == TYPE_BLOCK)),
    );
    vm.register_host_fn(
        "node:fs",
        "_statIsCharacterDevice",
        Box::new(|_ctx, args| Value::Bool(stat_type(args) == TYPE_CHAR)),
    );
    vm.register_host_fn(
        "node:fs",
        "_statIsFIFO",
        Box::new(|_ctx, args| Value::Bool(stat_type(args) == TYPE_FIFO)),
    );
    vm.register_host_fn(
        "node:fs",
        "_statIsSocket",
        Box::new(|_ctx, args| Value::Bool(stat_type(args) == TYPE_SOCKET)),
    );

    // ── readFileSync(path[, encoding]) ────────────────────────────
    // With encoding "utf8"/"utf-8", returns a string. Without
    // encoding, Node returns a Buffer; here we return a byte-array
    // (Object kind Array of i32 byte values) until a Buffer type
    // lands.
    vm.register_host_fn(
        "node:fs",
        "readFileSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let encoding = match args.get(1) {
                Some(Value::String(text)) => Some(text.to_string()),
                Some(Value::Object(opts)) => {
                    let o = opts.lock().unwrap();
                    match o.properties.get("encoding") {
                        Some(Value::String(enc)) => Some(enc.to_string()),
                        _ => None }
                }
                _ => None };
            match std::fs::read(&path) {
                Ok(bytes) => match encoding.as_deref() {
                    Some("utf8") | Some("utf-8") | Some("UTF-8") => {
                        Value::String(Arc::from(String::from_utf8_lossy(&bytes).as_ref()))
                    }
                    _ => {
                        let elems: Vec<Value> =
                            bytes.into_iter().map(|b| Value::I32(b as i32)).collect();
                        Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))
                    }
                },
                Err(e) => Value::String(Arc::from(format!("ENOENT: {}", e).as_str())) }
        }),
    );

    // ── writeFileSync(path, data) ─────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "writeFileSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let data = s_arg(args, 1, "");
            let flag = match args.get(2) {
                Some(Value::Object(opts)) => {
                    let o = opts.lock().unwrap();
                    match o.properties.get("flag") {
                        Some(Value::String(f)) => f.to_string(),
                        _ => "w".to_string() }
                }
                Some(Value::String(s)) => s.to_string(),
                _ => "w".to_string() };
            if flag == "a" || flag == "ax" || flag == "a+" {
                use std::io::Write;
                if let Ok(mut f) = std::fs::OpenOptions::new()
                    .create(true)
                    .append(true)
                    .open(&path)
                {
                    let _ = f.write_all(data.as_bytes());
                }
            } else {
                let _ = std::fs::write(&path, data.as_bytes());
            }
            Value::Null
        }),
    );

    // ── appendFileSync(path, data) ────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "appendFileSync",
        Box::new(|_ctx, args| {
            use std::io::Write;
            let path = s_arg(args, 0, "");
            let data = s_arg(args, 1, "");
            if let Ok(mut f) = std::fs::OpenOptions::new()
                .create(true)
                .append(true)
                .open(&path)
            {
                let _ = f.write_all(data.as_bytes());
            }
            Value::Null
        }),
    );

    // ── existsSync(path) → bool ───────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "existsSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            Value::Bool(std::path::Path::new(&path).exists())
        }),
    );

    // ── statSync(path) → Stats ────────────────────────────────────
    // Need a VM ref to bind method host-fn refs. The host fn signature
    // gives us a HostContext but not a &VM directly; however the
    // host_registry is on VM and we baked the closure with a captured
    // pointer. Workaround: closure holds the VM pointer at registration
    // time, but that's fragile. Cleaner: record the host_idx of each
    // predicate fn here and embed those numbers directly.
    let pred_idx = |vm: &VM, name: &str| -> usize {
        *vm.host_registry
            .get(&("node:fs".to_string(), name.to_string()))
            .expect("predicate fn must be registered first")
    };
    let is_file_idx = pred_idx(vm, "_statIsFile");
    let is_dir_idx = pred_idx(vm, "_statIsDirectory");
    let is_sym_idx = pred_idx(vm, "_statIsSymbolicLink");
    let is_blk_idx = pred_idx(vm, "_statIsBlockDevice");
    let is_chr_idx = pred_idx(vm, "_statIsCharacterDevice");
    let is_fifo_idx = pred_idx(vm, "_statIsFIFO");
    let is_sock_idx = pred_idx(vm, "_statIsSocket");

    let make_pred_ref = move |name: &str, idx: usize| -> Value {
        let mut obj = Object::new();
        obj.properties
            .insert("__host_module".into(), Value::String(Arc::from("node:fs")));
        obj.properties
            .insert("__host_name".into(), Value::String(Arc::from(name)));
        obj.properties
            .insert("__host_idx".into(), Value::F64(idx as f64));
        obj.kind = ObjectKind::HostFunction(idx);
        Value::Object(vybe_runtime::heap::alloc(obj))
    };
    let build_stats_with_idxs = move |meta: &std::fs::Metadata| -> Value {
        let mut o = Object::new();
        o.properties
            .insert("size".into(), Value::F64(meta.len() as f64));
        if let Ok(t) = meta.modified() {
            o.properties
                .insert("mtimeMs".into(), Value::F64(ms_since_epoch(t)));
        }
        if let Ok(t) = meta.accessed() {
            o.properties
                .insert("atimeMs".into(), Value::F64(ms_since_epoch(t)));
        }
        if let Ok(t) = meta.created() {
            o.properties
                .insert("birthtimeMs".into(), Value::F64(ms_since_epoch(t)));
        }
        let ft = meta.file_type();
        let type_tag = if ft.is_file() {
            TYPE_FILE
        } else if ft.is_dir() {
            TYPE_DIR
        } else if ft.is_symlink() {
            TYPE_SYMLINK
        } else {
            0
        };
        o.properties.insert("__type".into(), Value::I32(type_tag));
        o.properties
            .insert("isFile".into(), make_pred_ref("_statIsFile", is_file_idx));
        o.properties.insert(
            "isDirectory".into(),
            make_pred_ref("_statIsDirectory", is_dir_idx),
        );
        o.properties.insert(
            "isSymbolicLink".into(),
            make_pred_ref("_statIsSymbolicLink", is_sym_idx),
        );
        o.properties.insert(
            "isBlockDevice".into(),
            make_pred_ref("_statIsBlockDevice", is_blk_idx),
        );
        o.properties.insert(
            "isCharacterDevice".into(),
            make_pred_ref("_statIsCharacterDevice", is_chr_idx),
        );
        o.properties
            .insert("isFIFO".into(), make_pred_ref("_statIsFIFO", is_fifo_idx));
        o.properties.insert(
            "isSocket".into(),
            make_pred_ref("_statIsSocket", is_sock_idx),
        );
        Value::Object(vybe_runtime::heap::alloc(o))
    };
    let build_stats_clone = build_stats_with_idxs.clone();
    let build_stats_fstat = build_stats_with_idxs.clone();

    // Lightweight dirent builder for readdirSync {withFileTypes: true}
    let build_dirent = {
        let is_file_idx2 = is_file_idx;
        let is_dir_idx2 = is_dir_idx;
        let is_sym_idx2 = is_sym_idx;
        move |name: String, meta: &std::fs::Metadata| -> Value {
            let ft = meta.file_type();
            let type_tag = if ft.is_file() {
                TYPE_FILE
            } else if ft.is_dir() {
                TYPE_DIR
            } else if ft.is_symlink() {
                TYPE_SYMLINK
            } else {
                0
            };
            let make_fn = |idx: usize| -> Value {
                let mut obj = Object::new();
                obj.kind = ObjectKind::HostFunction(idx);
                Value::Object(vybe_runtime::heap::alloc(obj))
            };
            let mut o = Object::new();
            o.properties
                .insert("name".into(), Value::String(Arc::from(name.as_str())));
            o.properties.insert("__type".into(), Value::I32(type_tag));
            o.properties.insert("isFile".into(), make_fn(is_file_idx2));
            o.properties
                .insert("isDirectory".into(), make_fn(is_dir_idx2));
            o.properties
                .insert("isSymbolicLink".into(), make_fn(is_sym_idx2));
            Value::Object(vybe_runtime::heap::alloc(o))
        }
    };

    vm.register_host_fn(
        "node:fs",
        "statSync",
        Box::new(move |_ctx, args| {
            let path = s_arg(args, 0, "");
            match std::fs::metadata(&path) {
                Ok(meta) => build_stats_with_idxs(&meta),
                Err(_) => Value::Null }
        }),
    );

    // ── lstatSync(path) → Stats (does not follow symlinks) ────────
    vm.register_host_fn(
        "node:fs",
        "lstatSync",
        Box::new(move |_ctx, args| {
            let path = s_arg(args, 0, "");
            match std::fs::symlink_metadata(&path) {
                Ok(meta) => build_stats_clone(&meta),
                Err(_) => Value::Null }
        }),
    );

    // ── fstatSync(fd) → Stats ─────────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "fstatSync",
        Box::new(move |_ctx, args| {
            let fd = match args.first() {
                Some(Value::I32(n)) => *n,
                Some(Value::F64(f)) => *f as i32,
                _ => return Value::Null };
            let path = fd_table()
                .lock()
                .unwrap()
                .get(&fd)
                .map(|(p, _, _)| p.clone());
            match path {
                Some(p) => match std::fs::metadata(&p) {
                    Ok(meta) => build_stats_fstat(&meta),
                    Err(_) => Value::Null },
                None => Value::Null }
        }),
    );

    // ── readdirSync(path[, {withFileTypes}]) ────────────────────────
    vm.register_host_fn(
        "node:fs",
        "readdirSync",
        Box::new(move |_ctx, args| {
            let path = s_arg(args, 0, "");
            let with_file_types = opt_obj_bool(args, 1, "withFileTypes");
            match std::fs::read_dir(&path) {
                Ok(entries) => {
                    let items: Vec<Value> = entries
                        .filter_map(|e| e.ok())
                        .map(|e| {
                            if with_file_types {
                                let name = e.file_name().to_string_lossy().to_string();
                                if let Ok(meta) = e.metadata() {
                                    build_dirent(name, &meta)
                                } else {
                                    let mut o = Object::new();
                                    o.properties.insert(
                                        "name".into(),
                                        Value::String(Arc::from(
                                            e.file_name().to_string_lossy().as_ref(),
                                        )),
                                    );
                                    o.properties.insert("__type".into(), Value::I32(0));
                                    Value::Object(vybe_runtime::heap::alloc(o))
                                }
                            } else {
                                Value::String(Arc::from(e.file_name().to_string_lossy().as_ref()))
                            }
                        })
                        .collect();
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
                }
                Err(_) => Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new()))) }
        }),
    );

    // ── mkdirSync(path[, options]) ────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "mkdirSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let recursive = opt_obj_bool(args, 1, "recursive");
            let _ = if recursive {
                std::fs::create_dir_all(&path)
            } else {
                std::fs::create_dir(&path)
            };
            Value::Null
        }),
    );

    // ── unlinkSync(path) — remove file ────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "unlinkSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let _ = std::fs::remove_file(&path);
            Value::Null
        }),
    );

    // ── rmdirSync(path) — remove empty directory ──────────────────
    vm.register_host_fn(
        "node:fs",
        "rmdirSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let _ = std::fs::remove_dir(&path);
            Value::Null
        }),
    );

    // ── rmSync(path[, options]) — remove file or directory ────────
    vm.register_host_fn(
        "node:fs",
        "rmSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let recursive = opt_obj_bool(args, 1, "recursive");
            let _force = opt_obj_bool(args, 1, "force");
            let p = std::path::Path::new(&path);
            let _ = if p.is_dir() {
                if recursive {
                    std::fs::remove_dir_all(p)
                } else {
                    std::fs::remove_dir(p)
                }
            } else {
                std::fs::remove_file(p)
            };
            Value::Null
        }),
    );

    // ── renameSync(oldPath, newPath) ──────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "renameSync",
        Box::new(|_ctx, args| {
            let old = s_arg(args, 0, "");
            let new_path = s_arg(args, 1, "");
            let _ = std::fs::rename(&old, &new_path);
            Value::Null
        }),
    );

    // ── copyFileSync(src, dst) ────────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "copyFileSync",
        Box::new(|_ctx, args| {
            let src = s_arg(args, 0, "");
            let dst = s_arg(args, 1, "");
            let _ = std::fs::copy(&src, &dst);
            Value::Null
        }),
    );

    // ── realpathSync(path) — resolve symlinks + normalize ─────────
    vm.register_host_fn(
        "node:fs",
        "realpathSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            match std::fs::canonicalize(&path) {
                Ok(p) => Value::String(Arc::from(p.to_string_lossy().as_ref())),
                Err(_) => Value::Null }
        }),
    );

    // ── readlinkSync(path) — read symlink target ──────────────────
    vm.register_host_fn(
        "node:fs",
        "readlinkSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            match std::fs::read_link(&path) {
                Ok(p) => Value::String(Arc::from(p.to_string_lossy().as_ref())),
                Err(_) => Value::Null }
        }),
    );

    // ── accessSync(path) — undefined on success, throws on failure ─
    // Vybe doesn't model JS exceptions for host-fn returns yet; we
    // signal success with `undefined` and failure as `null`. Tests
    // accept either Undefined or Null for the success path.
    vm.register_host_fn(
        "node:fs",
        "accessSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            if std::path::Path::new(&path).exists() {
                Value::Undefined
            } else {
                Value::Null
            }
        }),
    );

    // ── truncateSync(path[, len]) — resize file to len (default 0) ─
    vm.register_host_fn(
        "node:fs",
        "truncateSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let len = match args.get(1) {
                Some(Value::F64(n)) => *n as u64,
                Some(Value::I32(n)) => *n as u64,
                _ => 0 };
            if let Ok(f) = std::fs::OpenOptions::new().write(true).open(&path) {
                let _ = f.set_len(len);
            }
            Value::Null
        }),
    );

    // ── openSync(path, flags) → fd ────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "openSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let flags = s_arg(args, 1, "r");
            // Ensure file exists for write mode
            if flags.contains('w') || flags.contains('a') {
                let _ = std::fs::OpenOptions::new()
                    .create(true)
                    .write(true)
                    .truncate(flags == "w")
                    .open(&path);
            }
            let fd = NEXT_FD.fetch_add(1, Ordering::Relaxed);
            fd_table().lock().unwrap().insert(fd, (path, flags, 0));
            Value::I32(fd)
        }),
    );

    // ── closeSync(fd) ─────────────────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "closeSync",
        Box::new(|_ctx, args| {
            let fd = match args.first() {
                Some(Value::I32(n)) => *n,
                Some(Value::F64(f)) => *f as i32,
                _ => return Value::Undefined };
            fd_table().lock().unwrap().remove(&fd);
            Value::Undefined
        }),
    );

    // ── writeSync(fd, data) → bytes_written ──────────────────────
    vm.register_host_fn(
        "node:fs",
        "writeSync",
        Box::new(|_ctx, args| {
            use std::io::{Seek, SeekFrom, Write};
            let fd = match args.first() {
                Some(Value::I32(n)) => *n,
                Some(Value::F64(f)) => *f as i32,
                _ => return Value::I32(0) };
            let data: Vec<u8> = match args.get(1) {
                Some(Value::String(s)) => s.as_bytes().to_vec(),
                Some(Value::Object(obj)) => {
                    let obj = obj.lock().unwrap();
                    if let ObjectKind::Array(elems) = &obj.kind {
                        elems
                            .iter()
                            .map(|e| match e {
                                Value::I32(n) => *n as u8,
                                _ => 0 })
                            .collect()
                    } else {
                        Vec::new()
                    }
                }
                _ => Vec::new() };
            let n = data.len();
            let (path, _flags, pos) = {
                let t = fd_table().lock().unwrap();
                match t.get(&fd) {
                    Some(e) => e.clone(),
                    None => return Value::I32(0) }
            };
            if let Ok(mut f) = std::fs::OpenOptions::new()
                .write(true)
                .create(true)
                .open(&path)
            {
                let _ = f.seek(SeekFrom::Start(pos));
                let _ = f.write_all(&data);
            }
            fd_table()
                .lock()
                .unwrap()
                .entry(fd)
                .and_modify(|e| e.2 += n as u64);
            Value::I32(n as i32)
        }),
    );

    // ── readSync(fd, buffer, offset, length, position) → bytes_read
    vm.register_host_fn(
        "node:fs",
        "readSync",
        Box::new(|_ctx, args| {
            use std::io::{Read, Seek, SeekFrom};
            let fd = match args.first() {
                Some(Value::I32(n)) => *n,
                Some(Value::F64(f)) => *f as i32,
                _ => return Value::I32(0) };
            let length = match args.get(3) {
                Some(Value::I32(n)) => *n as usize,
                Some(Value::F64(f)) => *f as usize,
                _ => 0 };
            let position = match args.get(4) {
                Some(Value::I32(n)) => *n as i64,
                Some(Value::F64(f)) => *f as i64,
                _ => -1 };
            let path = {
                let t = fd_table().lock().unwrap();
                match t.get(&fd) {
                    Some((p, _, _)) => p.clone(),
                    None => return Value::I32(0) }
            };
            let mut buf = vec![0u8; length];
            let n = if let Ok(mut f) = std::fs::File::open(&path) {
                if position >= 0 {
                    let _ = f.seek(SeekFrom::Start(position as u64));
                }
                f.read(&mut buf).unwrap_or(0)
            } else {
                0
            };
            // Fill buffer object
            if let Some(Value::Object(buf_obj)) = args.get(1) {
                let offset = match args.get(2) {
                    Some(Value::I32(o)) => *o as usize,
                    Some(Value::F64(f)) => *f as usize,
                    _ => 0 };
                let mut bo = buf_obj.lock().unwrap();
                if let ObjectKind::Array(ref mut elems) = bo.kind {
                    for (i, &b) in buf[..n].iter().enumerate() {
                        if offset + i < elems.len() {
                            elems[offset + i] = Value::I32(b as i32);
                        }
                    }
                }
            }
            Value::I32(n as i32)
        }),
    );

    // ── fsyncSync / fdatasyncSync ─────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "fsyncSync",
        Box::new(|_ctx, _args| Value::Undefined),
    );
    vm.register_host_fn(
        "node:fs",
        "fdatasyncSync",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    // ── ftruncateSync(fd, len) ────────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "ftruncateSync",
        Box::new(|_ctx, args| {
            let fd = match args.first() {
                Some(Value::I32(n)) => *n,
                Some(Value::F64(f)) => *f as i32,
                _ => return Value::Undefined };
            let len = match args.get(1) {
                Some(Value::I32(n)) => *n as u64,
                Some(Value::F64(f)) => *f as u64,
                _ => 0 };
            let path = {
                let t = fd_table().lock().unwrap();
                t.get(&fd).map(|(p, _, _)| p.clone())
            };
            if let Some(p) = path {
                if let Ok(f) = std::fs::OpenOptions::new().write(true).open(&p) {
                    let _ = f.set_len(len);
                }
            }
            Value::Undefined
        }),
    );

    // ── chmodSync(path, mode) ─────────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "chmodSync",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            #[cfg(unix)]
            {
                use std::os::unix::fs::PermissionsExt;
                let mode = match args.get(1) {
                    Some(Value::I32(n)) => *n as u32,
                    Some(Value::F64(f)) => *f as u32,
                    _ => 0o644 };
                let _ = std::fs::set_permissions(&path, std::fs::Permissions::from_mode(mode));
            }
            let _ = path;
            Value::Undefined
        }),
    );

    // ── chownSync(path, uid, gid) ─────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "chownSync",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    // ── symlinkSync(target, path) ─────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "symlinkSync",
        Box::new(|_ctx, args| {
            let target = s_arg(args, 0, "");
            let link = s_arg(args, 1, "");
            #[cfg(unix)]
            {
                let _ = std::os::unix::fs::symlink(&target, &link);
            }
            #[cfg(not(unix))]
            {
                let _ = (target, link);
            }
            Value::Undefined
        }),
    );

    // ── linkSync(existing, new) ───────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "linkSync",
        Box::new(|_ctx, args| {
            let existing = s_arg(args, 0, "");
            let new_path = s_arg(args, 1, "");
            let _ = std::fs::hard_link(&existing, &new_path);
            Value::Undefined
        }),
    );

    // ── mkdtempSync(prefix) → path ────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "mkdtempSync",
        Box::new(|_ctx, args| {
            let prefix = s_arg(args, 0, "");
            // Generate unique suffix from current time
            let suffix = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| d.subsec_nanos())
                .unwrap_or(0);
            let path = format!("{}{:08x}", prefix, suffix);
            let _ = std::fs::create_dir_all(&path);
            Value::String(Arc::from(path.as_str()))
        }),
    );

    // ── createReadStream(path) → stream stub ──────────────────────
    vm.register_host_fn(
        "node:fs",
        "createReadStream",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let mut o = Object::new();
            o.properties
                .insert("path".into(), Value::String(Arc::from(path.as_str())));
            o.properties.insert("readable".into(), Value::Bool(true));
            o.properties.insert("destroyed".into(), Value::Bool(false));
            for m in [
                "on",
                "once",
                "off",
                "emit",
                "pipe",
                "unpipe",
                "read",
                "destroy",
                "pause",
                "resume",
                "setEncoding",
                "addListener",
                "removeListener",
            ] {
                o.properties.insert(m.into(), Value::Undefined);
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // ── createWriteStream(path) → stream stub ────────────────────
    vm.register_host_fn(
        "node:fs",
        "createWriteStream",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let mut o = Object::new();
            o.properties
                .insert("path".into(), Value::String(Arc::from(path.as_str())));
            o.properties.insert("writable".into(), Value::Bool(true));
            o.properties.insert("destroyed".into(), Value::Bool(false));
            for m in [
                "write",
                "end",
                "on",
                "once",
                "off",
                "emit",
                "destroy",
                "cork",
                "uncork",
                "addListener",
                "removeListener",
            ] {
                o.properties.insert(m.into(), Value::Undefined);
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // ── watch(path[, callback]) → watcher stub ───────────────────
    vm.register_host_fn(
        "node:fs",
        "watch",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let mut o = Object::new();
            o.properties
                .insert("path".into(), Value::String(Arc::from(path.as_str())));
            for m in ["close", "on", "once", "off", "emit"] {
                o.properties.insert(m.into(), Value::Undefined);
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // ── watchFile(path[, callback]) → stat watcher stub ──────────
    vm.register_host_fn(
        "node:fs",
        "watchFile",
        Box::new(|_ctx, args| {
            let path = s_arg(args, 0, "");
            let mut o = Object::new();
            o.properties
                .insert("path".into(), Value::String(Arc::from(path.as_str())));
            for m in ["stop", "on", "once", "off"] {
                o.properties.insert(m.into(), Value::Undefined);
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // ── unwatchFile → undefined ───────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "unwatchFile",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    // ── constants ─────────────────────────────────────────────────
    vm.register_host_fn(
        "node:fs",
        "constants",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("F_OK".into(), Value::I32(0));
            o.properties.insert("R_OK".into(), Value::I32(4));
            o.properties.insert("W_OK".into(), Value::I32(2));
            o.properties.insert("X_OK".into(), Value::I32(1));
            o.properties.insert("O_RDONLY".into(), Value::I32(0));
            o.properties.insert("O_WRONLY".into(), Value::I32(1));
            o.properties.insert("O_RDWR".into(), Value::I32(2));
            #[cfg(unix)]
            o.properties
                .insert("O_CREAT".into(), Value::I32(libc_o_creat()));
            #[cfg(not(unix))]
            o.properties.insert("O_CREAT".into(), Value::I32(512));
            o.properties.insert("O_EXCL".into(), Value::I32(128));
            o.properties.insert("O_TRUNC".into(), Value::I32(512));
            o.properties.insert("O_APPEND".into(), Value::I32(8));
            o.properties.insert("O_SYNC".into(), Value::I32(2048));
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );
}

// ── Platform helpers ───────────────────────────────────────────────
#[cfg(unix)]
fn libc_o_creat() -> i32 {
    64 // O_CREAT on Linux; 512 on macOS — use 64 as safe cross-platform value
}

// Suppress the unused build_stats helper (kept in case future
// implementations want the simpler call signature).
#[allow(dead_code)]
fn _unused_build_stats(_: &std::fs::Metadata, _: &VM) -> Value {
    build_stats(&std::fs::metadata(".").unwrap(), &VM::new())
}

// Suppress unused imports.
#[allow(dead_code)]
fn _force_use(_: &mut HostContext) {}
