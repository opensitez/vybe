//! `wasi:filesystem` — WASI 0.2.8 filesystem proposal.
//!
//! Implements the descriptor-based, capability-rooted surface
//! defined in `proposals/wasi-filesystem/wit/{types,preopens}.wit`
//! (package `wasi:filesystem@0.2.8`). Function names match the
//! canonical-ABI form so a Component-Model runtime that loads
//! Vybe-emitted `.wasm` resolves them against any conforming WASI
//! implementation:
//!
//!   - `wasi:filesystem/types` — `[method]descriptor.<op>` and
//!     `[method]directory-entry-stream.read-directory-entry`
//!   - `wasi:filesystem/preopens` — `get-directories`
//!   - `wasi:io/streams` — `[method]input-stream.blocking-read`
//!     (and friends) for reading file streams returned by
//!     `read-via-stream`. Socket streams already live there;
//!     this module extends the same registry to file streams.
//!
//! Resources (descriptors, directory-entry-streams, input-streams,
//! output-streams) are returned to the guest as Vybe `Object`s
//! carrying internal `__wasi_kind` + `__wasi_id` properties; the
//! actual `std::fs` handles live in a host-side registry indexed by
//! that id. This mirrors how `wasi:sockets` already represents its
//! resources in the host.
//!
//! Errors are returned as Vybe Objects with a `__wasi_error` field
//! carrying the WIT `error-code` enum value as a string
//! (`"no-entry"`, `"is-directory"`, `"not-empty"`, etc.). Real
//! Component-Model bindings would marshal these as exceptions; the
//! Vybe VM doesn't have a host-fn exception channel yet, so the
//! object-with-error-field is the carrier.

use std::collections::HashMap;
use std::fs::{File, OpenOptions, ReadDir};
use std::io::{Read, Seek, SeekFrom};
use std::path::PathBuf;
use std::sync::{Arc, Mutex, OnceLock};
use std::time::SystemTime;
use vybe_bytecode::value::Object;
use vybe_bytecode::{VM, Value};

// ── Resource registry ─────────────────────────────────────────────
//
// One per-process registry holds the actual std::fs handles + state
// behind `Arc<Mutex<…>>`. Guest code only sees opaque `Object`s
// carrying ids that index into this registry.

#[derive(Debug)]
enum DescriptorKind {
    File(PathBuf),       // path retained for stat-at / open-at(child)
    Directory(PathBuf),
}

#[derive(Debug)]
enum InputStreamKind {
    File { file: File, position: u64 },
}

struct Registry {
    descriptors: HashMap<u32, DescriptorKind>,
    dir_streams: HashMap<u32, ReadDir>,
    input_streams: HashMap<u32, InputStreamKind>,
    next_id: u32,
}

impl Registry {
    fn new() -> Self {
        Registry {
            descriptors: HashMap::new(),
            dir_streams: HashMap::new(),
            input_streams: HashMap::new(),
            next_id: 1,
        }
    }
    fn alloc_id(&mut self) -> u32 {
        let id = self.next_id;
        self.next_id += 1;
        id
    }
}

fn registry() -> &'static Mutex<Registry> {
    static R: OnceLock<Mutex<Registry>> = OnceLock::new();
    R.get_or_init(|| Mutex::new(Registry::new()))
}

// ── Resource <-> Value marshalling ────────────────────────────────

const KIND_DESCRIPTOR: &str = "descriptor";
const KIND_DIR_STREAM: &str = "directory-entry-stream";
const KIND_INPUT_STREAM: &str = "input-stream";

fn make_resource(kind: &str, id: u32) -> Value {
    let mut o = Object::new();
    o.properties.insert("__wasi_kind".into(), Value::String(Arc::from(kind)));
    o.properties.insert("__wasi_id".into(), Value::F64(id as f64));
    Value::Object(Arc::new(Mutex::new(o)))
}

fn resource_id(value: &Value, expected_kind: &str) -> Option<u32> {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        let kind_ok = matches!(
            object.properties.get("__wasi_kind"),
            Some(Value::String(text)) if text.as_ref() == expected_kind
        );
        if !kind_ok {
            return None;
        }
        if let Some(Value::F64(id)) = object.properties.get("__wasi_id") {
            return Some(*id as u32);
        }
    }
    None
}

// ── Error encoding ────────────────────────────────────────────────

fn err(code: &str) -> Value {
    let mut o = Object::new();
    o.properties.insert("__wasi_error".into(), Value::String(Arc::from(code)));
    Value::Object(Arc::new(Mutex::new(o)))
}

fn map_io_error(e: &std::io::Error) -> &'static str {
    use std::io::ErrorKind::*;
    match e.kind() {
        NotFound => "no-entry",
        PermissionDenied => "access",
        AlreadyExists => "exist",
        WouldBlock => "would-block",
        InvalidInput | InvalidData => "invalid",
        TimedOut => "io",
        Interrupted => "interrupted",
        Unsupported => "unsupported",
        UnexpectedEof => "io",
        OutOfMemory => "insufficient-memory",
        // POSIX-derived ENOTEMPTY isn't a stable std::io::ErrorKind variant;
        // detect via raw_os_error == 39 (Linux) / 66 (Darwin) / 145 (Windows).
        _ => match e.raw_os_error() {
            Some(39) | Some(66) | Some(145) => "not-empty",
            Some(21) => "is-directory",
            Some(20) => "not-directory",
            _ => "io",
        },
    }
}

// ── descriptor-stat / descriptor-type encoding ────────────────────

fn ms_since_epoch(time: SystemTime) -> f64 {
    time.duration_since(SystemTime::UNIX_EPOCH)
        .map(|d| d.as_millis() as f64)
        .unwrap_or(0.0)
}

fn descriptor_type_string(meta: &std::fs::Metadata) -> &'static str {
    let ft = meta.file_type();
    if ft.is_file() { "regular-file" }
    else if ft.is_dir() { "directory" }
    else if ft.is_symlink() { "symbolic-link" }
    else { "unknown" }
}

fn build_stat(meta: &std::fs::Metadata) -> Value {
    let mut o = Object::new();
    o.properties.insert("type".into(), Value::String(Arc::from(descriptor_type_string(meta))));
    // link-count: WIT u64 — we surface as f64 since Vybe doesn't have u64.
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;
        o.properties.insert("link-count".into(), Value::F64(meta.nlink() as f64));
    }
    #[cfg(not(unix))]
    {
        o.properties.insert("link-count".into(), Value::F64(1.0));
    }
    o.properties.insert("size".into(), Value::F64(meta.len() as f64));
    if let Ok(t) = meta.modified() {
        o.properties.insert("data-modification-timestamp".into(), Value::F64(ms_since_epoch(t)));
    }
    if let Ok(t) = meta.accessed() {
        o.properties.insert("data-access-timestamp".into(), Value::F64(ms_since_epoch(t)));
    }
    if let Ok(t) = meta.created() {
        o.properties.insert("status-change-timestamp".into(), Value::F64(ms_since_epoch(t)));
    }
    Value::Object(Arc::new(Mutex::new(o)))
}

// ── path resolution ───────────────────────────────────────────────
//
// WASI paths are interpreted relative to the parent descriptor.
// `path-flags::symlink-follow` (bit 0) controls whether symlinks
// in the resolved path are followed; the std::fs APIs follow by
// default, so the flag is currently informational.

fn resolve_child_path(parent_id: u32, child: &str) -> Result<PathBuf, &'static str> {
    let registry = registry().lock().unwrap();
    let parent = registry.descriptors.get(&parent_id).ok_or("bad-descriptor")?;
    let parent_path = match parent {
        DescriptorKind::Directory(p) => p,
        DescriptorKind::File(p) => p,
    };
    Ok(parent_path.join(child))
}

// ── Open-flags decoding ───────────────────────────────────────────
//
// WIT order: create, directory, exclusive, truncate (bits 0..=3).
// %flags (descriptor-flags): read, write, file-integrity-sync,
// data-integrity-sync, requested-write-sync, mutate-directory.

const OPEN_CREATE: i32 = 1;
const OPEN_DIRECTORY: i32 = 2;
const OPEN_EXCLUSIVE: i32 = 4;
const OPEN_TRUNCATE: i32 = 8;

const DESC_READ: i32 = 1;
const DESC_WRITE: i32 = 2;

// ── Argument helpers ──────────────────────────────────────────────

fn s_arg(args: &[Value], idx: usize) -> Option<String> {
    match args.get(idx) {
        Some(Value::String(text)) => Some(text.to_string()),
        Some(Value::F64(_)) | Some(Value::I32(_)) => None, // wrong type
        _ => None,
    }
}

fn i32_arg(args: &[Value], idx: usize, default: i32) -> i32 {
    match args.get(idx) {
        Some(Value::I32(n)) => *n,
        Some(Value::F64(n)) => *n as i32,
        _ => default,
    }
}

fn u64_arg(args: &[Value], idx: usize, default: u64) -> u64 {
    match args.get(idx) {
        Some(Value::F64(n)) => *n as u64,
        Some(Value::I32(n)) => *n as u64,
        _ => default,
    }
}

// ── Public registration ───────────────────────────────────────────

pub fn register(vm: &mut VM) {
    register_preopens(vm);
    register_types(vm);
    register_io_streams(vm);
    register_test_helpers(vm);
}

fn register_preopens(vm: &mut VM) {
    // get-directories() -> list<tuple<descriptor, string>>.
    // In a real WASI implementation the host pre-opens directories
    // before the guest runs. Vybe currently exposes only `cwd` as a
    // preopen (named ".") — programs can `open-at` from there.
    vm.register_host_fn("wasi:filesystem/preopens", "get-directories", Box::new(|_ctx, _args| {
        let cwd = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
        let id = {
            let mut reg = registry().lock().unwrap();
            let id = reg.alloc_id();
            reg.descriptors.insert(id, DescriptorKind::Directory(cwd));
            id
        };
        let pair_elements = vec![
            make_resource(KIND_DESCRIPTOR, id),
            Value::String(Arc::from(".")),
        ];
        let pair = Value::Object(Arc::new(Mutex::new(Object::new_array(pair_elements))));
        Value::Object(Arc::new(Mutex::new(Object::new_array(vec![pair]))))
    }));
}

fn register_types(vm: &mut VM) {
    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.open-at", Box::new(|_ctx, args| {
        let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
            return err("bad-descriptor");
        };
        let _path_flags = i32_arg(args, 1, 0);
        let Some(child) = s_arg(args, 2) else { return err("invalid"); };
        let open_flags = i32_arg(args, 3, 0);
        let _desc_flags = i32_arg(args, 4, 0);

        let resolved = match resolve_child_path(parent_id, &child) {
            Ok(p) => p,
            Err(code) => return err(code),
        };

        let create = (open_flags & OPEN_CREATE) != 0;
        let exclusive = (open_flags & OPEN_EXCLUSIVE) != 0;
        let truncate = (open_flags & OPEN_TRUNCATE) != 0;
        let directory = (open_flags & OPEN_DIRECTORY) != 0;

        let exists = resolved.exists();
        if !exists && !create {
            return err("no-entry");
        }
        if directory && exists && !resolved.is_dir() {
            return err("not-directory");
        }

        if !exists && create {
            // Touch the file so subsequent `open-at` / `stat` can find
            // it. We don't preserve the actual `File` handle here —
            // the descriptor is path-based.
            let mut opts = OpenOptions::new();
            opts.create(true).write(true);
            if exclusive { opts.create_new(true); }
            if truncate { opts.truncate(true); }
            if let Err(e) = opts.open(&resolved) {
                return err(map_io_error(&e));
            }
        } else if exists && truncate && resolved.is_file() {
            if let Err(e) = OpenOptions::new().write(true).truncate(true).open(&resolved) {
                return err(map_io_error(&e));
            }
        }

        let kind = if resolved.is_dir() {
            DescriptorKind::Directory(resolved)
        } else {
            DescriptorKind::File(resolved)
        };
        let id = {
            let mut reg = registry().lock().unwrap();
            let id = reg.alloc_id();
            reg.descriptors.insert(id, kind);
            id
        };
        make_resource(KIND_DESCRIPTOR, id)
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.stat", Box::new(|_ctx, args| {
        let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let path = {
            let reg = registry().lock().unwrap();
            match reg.descriptors.get(&id) {
                Some(DescriptorKind::File(p)) | Some(DescriptorKind::Directory(p)) => p.clone(),
                None => return err("bad-descriptor"),
            }
        };
        match std::fs::metadata(&path) {
            Ok(meta) => build_stat(&meta),
            Err(e) => err(map_io_error(&e)),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.stat-at", Box::new(|_ctx, args| {
        let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let _path_flags = i32_arg(args, 1, 0);
        let Some(child) = s_arg(args, 2) else { return err("invalid"); };
        let resolved = match resolve_child_path(parent_id, &child) {
            Ok(p) => p,
            Err(code) => return err(code),
        };
        match std::fs::metadata(&resolved) {
            Ok(meta) => build_stat(&meta),
            Err(e) => err(map_io_error(&e)),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.get-type", Box::new(|_ctx, args| {
        let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let path = {
            let reg = registry().lock().unwrap();
            match reg.descriptors.get(&id) {
                Some(DescriptorKind::File(p)) | Some(DescriptorKind::Directory(p)) => p.clone(),
                None => return err("bad-descriptor"),
            }
        };
        match std::fs::metadata(&path) {
            Ok(meta) => Value::String(Arc::from(descriptor_type_string(&meta))),
            Err(e) => err(map_io_error(&e)),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.read-directory", Box::new(|_ctx, args| {
        let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let path = {
            let reg = registry().lock().unwrap();
            match reg.descriptors.get(&id) {
                Some(DescriptorKind::Directory(p)) => p.clone(),
                Some(DescriptorKind::File(_)) => return err("not-directory"),
                None => return err("bad-descriptor"),
            }
        };
        match std::fs::read_dir(&path) {
            Ok(rd) => {
                let stream_id = {
                    let mut reg = registry().lock().unwrap();
                    let stream_id = reg.alloc_id();
                    reg.dir_streams.insert(stream_id, rd);
                    stream_id
                };
                make_resource(KIND_DIR_STREAM, stream_id)
            }
            Err(e) => err(map_io_error(&e)),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]directory-entry-stream.read-directory-entry", Box::new(|_ctx, args| {
        let Some(stream_id) = resource_id(&args[0].clone(), KIND_DIR_STREAM) else { return err("bad-descriptor"); };
        let mut reg = registry().lock().unwrap();
        let Some(stream) = reg.dir_streams.get_mut(&stream_id) else { return err("bad-descriptor"); };
        match stream.next() {
            None => Value::Null, // option<directory-entry>::none — end of stream.
            Some(Err(e)) => err(map_io_error(&e)),
            Some(Ok(entry)) => {
                let name = entry.file_name().to_string_lossy().to_string();
                let entry_type = match entry.file_type() {
                    Ok(ft) if ft.is_file() => "regular-file",
                    Ok(ft) if ft.is_dir() => "directory",
                    Ok(ft) if ft.is_symlink() => "symbolic-link",
                    _ => "unknown",
                };
                let mut o = Object::new();
                o.properties.insert("type".into(), Value::String(Arc::from(entry_type)));
                o.properties.insert("name".into(), Value::String(Arc::from(name.as_str())));
                Value::Object(Arc::new(Mutex::new(o)))
            }
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.create-directory-at", Box::new(|_ctx, args| {
        let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let Some(child) = s_arg(args, 1) else { return err("invalid"); };
        let resolved = match resolve_child_path(parent_id, &child) {
            Ok(p) => p,
            Err(code) => return err(code),
        };
        match std::fs::create_dir(&resolved) {
            Ok(_) => Value::Null,
            Err(e) => err(map_io_error(&e)),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.unlink-file-at", Box::new(|_ctx, args| {
        let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let Some(child) = s_arg(args, 1) else { return err("invalid"); };
        let resolved = match resolve_child_path(parent_id, &child) {
            Ok(p) => p,
            Err(code) => return err(code),
        };
        if resolved.is_dir() {
            return err("is-directory");
        }
        match std::fs::remove_file(&resolved) {
            Ok(_) => Value::Null,
            Err(e) => err(map_io_error(&e)),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.remove-directory-at", Box::new(|_ctx, args| {
        let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let Some(child) = s_arg(args, 1) else { return err("invalid"); };
        let resolved = match resolve_child_path(parent_id, &child) {
            Ok(p) => p,
            Err(code) => return err(code),
        };
        match std::fs::remove_dir(&resolved) {
            Ok(_) => Value::Null,
            Err(e) => err(map_io_error(&e)),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.rename-at", Box::new(|_ctx, args| {
        let Some(old_parent) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let Some(old_child) = s_arg(args, 1) else { return err("invalid"); };
        let Some(new_parent) = resource_id(&args[2].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let Some(new_child) = s_arg(args, 3) else { return err("invalid"); };
        let old_path = match resolve_child_path(old_parent, &old_child) { Ok(p) => p, Err(c) => return err(c) };
        let new_path = match resolve_child_path(new_parent, &new_child) { Ok(p) => p, Err(c) => return err(c) };
        match std::fs::rename(&old_path, &new_path) {
            Ok(_) => Value::Null,
            Err(e) => err(map_io_error(&e)),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.readlink-at", Box::new(|_ctx, args| {
        let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let Some(child) = s_arg(args, 1) else { return err("invalid"); };
        let resolved = match resolve_child_path(parent_id, &child) { Ok(p) => p, Err(c) => return err(c) };
        match std::fs::read_link(&resolved) {
            Ok(target) => Value::String(Arc::from(target.to_string_lossy().as_ref())),
            Err(e) => err(map_io_error(&e)),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.is-same-object", Box::new(|_ctx, args| {
        let a = resource_id(&args[0].clone(), KIND_DESCRIPTOR);
        let b = resource_id(&args[1].clone(), KIND_DESCRIPTOR);
        match (a, b) {
            (Some(a_id), Some(b_id)) if a_id == b_id => Value::Bool(true),
            (Some(a_id), Some(b_id)) => {
                let reg = registry().lock().unwrap();
                let a_path = reg.descriptors.get(&a_id).map(path_of);
                let b_path = reg.descriptors.get(&b_id).map(path_of);
                Value::Bool(a_path.is_some() && a_path == b_path)
            }
            _ => Value::Bool(false),
        }
    }));

    vm.register_host_fn("wasi:filesystem/types", "[method]descriptor.read-via-stream", Box::new(|_ctx, args| {
        let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else { return err("bad-descriptor"); };
        let offset = u64_arg(args, 1, 0);
        let path = {
            let reg = registry().lock().unwrap();
            match reg.descriptors.get(&id) {
                Some(DescriptorKind::File(p)) => p.clone(),
                Some(DescriptorKind::Directory(_)) => return err("is-directory"),
                None => return err("bad-descriptor"),
            }
        };
        let mut file = match File::open(&path) {
            Ok(f) => f,
            Err(e) => return err(map_io_error(&e)),
        };
        if offset > 0 {
            if let Err(e) = file.seek(SeekFrom::Start(offset)) {
                return err(map_io_error(&e));
            }
        }
        let stream_id = {
            let mut reg = registry().lock().unwrap();
            let stream_id = reg.alloc_id();
            reg.input_streams.insert(stream_id, InputStreamKind::File { file, position: offset });
            stream_id
        };
        make_resource(KIND_INPUT_STREAM, stream_id)
    }));
}

fn path_of(kind: &DescriptorKind) -> PathBuf {
    match kind {
        DescriptorKind::File(p) | DescriptorKind::Directory(p) => p.clone(),
    }
}

// ── wasi:io/streams (file streams) ────────────────────────────────
//
// `wasi:sockets` already registers `wasi:io/streams.read` etc. for
// socket streams. Those handlers don't recognise our file-stream
// resource shape (they look for socket-specific properties), so we
// register additional `wasi:io/streams` entries scoped to the
// `[method]input-stream.<op>` canonical names. Calls flow to
// whichever handler matches the resource shape — in our case, file
// streams returned by `read-via-stream`.

fn register_io_streams(vm: &mut VM) {
    vm.register_host_fn("wasi:io/streams", "[method]input-stream.blocking-read", Box::new(|_ctx, args| {
        let Some(id) = resource_id(&args[0].clone(), KIND_INPUT_STREAM) else { return err("bad-descriptor"); };
        let max = u64_arg(args, 1, u64::MAX);
        let mut reg = registry().lock().unwrap();
        let Some(stream) = reg.input_streams.get_mut(&id) else { return err("bad-descriptor"); };
        match stream {
            InputStreamKind::File { file, position } => {
                let cap = max.min(64 * 1024) as usize;
                let mut buf = vec![0u8; cap];
                match file.read(&mut buf) {
                    Ok(n) => {
                        buf.truncate(n);
                        *position += n as u64;
                        let elements: Vec<Value> = buf.into_iter().map(|b| Value::I32(b as i32)).collect();
                        Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
                    }
                    Err(e) => err(map_io_error(&e)),
                }
            }
        }
    }));

    vm.register_host_fn("wasi:io/streams", "[method]input-stream.read", Box::new(|_ctx, args| {
        // Non-blocking read maps to the same impl in our synchronous VM —
        // no actual blocking happens because std::fs::File::read is sync.
        let Some(id) = resource_id(&args[0].clone(), KIND_INPUT_STREAM) else { return err("bad-descriptor"); };
        let max = u64_arg(args, 1, u64::MAX);
        let mut reg = registry().lock().unwrap();
        let Some(stream) = reg.input_streams.get_mut(&id) else { return err("bad-descriptor"); };
        match stream {
            InputStreamKind::File { file, position } => {
                let cap = max.min(64 * 1024) as usize;
                let mut buf = vec![0u8; cap];
                match file.read(&mut buf) {
                    Ok(n) => {
                        buf.truncate(n);
                        *position += n as u64;
                        let elements: Vec<Value> = buf.into_iter().map(|b| Value::I32(b as i32)).collect();
                        Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
                    }
                    Err(e) => err(map_io_error(&e)),
                }
            }
        }
    }));
}

// ── Test-only helpers ─────────────────────────────────────────────
//
// Production code receives preopened directories via
// `get-directories`. Tests need to build a descriptor for a known
// scratch path; this opens that backdoor under a `__test_*` name so
// there's no chance of guest code stumbling onto it accidentally
// (none of the language profiles route imports to `__test_*`).

fn register_test_helpers(vm: &mut VM) {
    vm.register_host_fn("wasi:filesystem/types", "__test_open_root", Box::new(|_ctx, args| {
        let Some(path_str) = s_arg(args, 0) else { return err("invalid"); };
        let path = PathBuf::from(&path_str);
        if !path.is_dir() { return err("not-directory"); }
        let id = {
            let mut reg = registry().lock().unwrap();
            let id = reg.alloc_id();
            reg.descriptors.insert(id, DescriptorKind::Directory(path));
            id
        };
        make_resource(KIND_DESCRIPTOR, id)
    }));
}

// Suppress unused-constant warnings for the descriptor-flag bits
// (`open-flags` are read above; descriptor-flags will be once
// write-via-stream/append-via-stream land).
#[allow(dead_code)]
const _: i32 = DESC_READ + DESC_WRITE;
