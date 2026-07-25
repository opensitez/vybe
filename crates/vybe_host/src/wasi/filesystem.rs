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
use std::fs::{File, FileTimes, OpenOptions, ReadDir};
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::PathBuf;
use std::sync::{Arc, Mutex, OnceLock};
use std::time::{Duration, SystemTime};
use vybe_bytecode::value::Object;
use vybe_bytecode::{VM, Value};

// ── Resource registry ─────────────────────────────────────────────
//
// One per-process registry holds the actual std::fs handles + state
// behind `Arc<Mutex<…>>`. Guest code only sees opaque `Object`s
// carrying ids that index into this registry.

#[derive(Debug)]
pub(super) enum DescriptorKind {
    File(PathBuf), // path retained for stat-at / open-at(child)
    Directory(PathBuf),
}

#[derive(Debug)]
pub(super) enum InputStreamKind {
    File { file: File, position: u64 },
    Buffer { data: Vec<u8>, position: usize },
}

#[derive(Debug)]
pub(super) enum OutputStreamKind {
    File { file: File },
    Append(PathBuf),
}

pub(super) struct Registry {
    pub(super) descriptors: HashMap<u32, DescriptorKind>,
    pub(super) dir_streams: HashMap<u32, ReadDir>,
    pub(super) input_streams: HashMap<u32, InputStreamKind>,
    pub(super) output_streams: HashMap<u32, OutputStreamKind>,
    next_id: u32,
}

impl Registry {
    fn new() -> Self {
        Registry {
            descriptors: HashMap::new(),
            dir_streams: HashMap::new(),
            input_streams: HashMap::new(),
            output_streams: HashMap::new(),
            next_id: 1,
        }
    }
    fn alloc_id(&mut self) -> u32 {
        let id = self.next_id;
        self.next_id += 1;
        id
    }
}

pub(super) fn registry() -> &'static Mutex<Registry> {
    static R: OnceLock<Mutex<Registry>> = OnceLock::new();
    R.get_or_init(|| Mutex::new(Registry::new()))
}

/// Register an in-memory byte buffer as a readable `input-stream` resource.
/// Returns the stream id; callers should wrap it as `make_resource(KIND_INPUT_STREAM, id)`.
/// Used by `wasi:http/types.[method]incoming-body.stream` to expose response body bytes
/// through the standard `[method]input-stream.blocking-read` interface.
pub fn register_buffer_stream(data: Vec<u8>) -> u32 {
    let mut reg = registry().lock().unwrap();
    let id = reg.alloc_id();
    reg.input_streams
        .insert(id, InputStreamKind::Buffer { data, position: 0 });
    id
}

// ── Resource <-> Value marshalling ────────────────────────────────

const KIND_DESCRIPTOR: &str = "descriptor";
const KIND_DIR_STREAM: &str = "directory-entry-stream";
pub(super) const KIND_INPUT_STREAM: &str = "input-stream";
pub(super) const KIND_OUTPUT_STREAM: &str = "output-stream";

fn make_resource(kind: &str, id: u32) -> Value {
    let mut o = Object::new();
    o.properties
        .insert("__wasi_kind".into(), Value::String(Arc::from(kind)));
    o.properties
        .insert("__wasi_id".into(), Value::F64(id as f64));
    Value::Object(vybe_bytecode::heap::alloc(o))
}

pub(super) fn resource_id(value: &Value, expected_kind: &str) -> Option<u32> {
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

pub(super) fn err(code: &str) -> Value {
    let mut o = Object::new();
    o.properties
        .insert("__wasi_error".into(), Value::String(Arc::from(code)));
    Value::Object(vybe_bytecode::heap::alloc(o))
}

pub(super) fn map_io_error(e: &std::io::Error) -> &'static str {
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
    if ft.is_file() {
        "regular-file"
    } else if ft.is_dir() {
        "directory"
    } else if ft.is_symlink() {
        "symbolic-link"
    } else {
        "unknown"
    }
}

fn build_stat(meta: &std::fs::Metadata) -> Value {
    let mut o = Object::new();
    o.properties.insert(
        "type".into(),
        Value::String(Arc::from(descriptor_type_string(meta))),
    );
    // link-count: WIT u64 — we surface as f64 since Vybe doesn't have u64.
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;
        o.properties
            .insert("link-count".into(), Value::F64(meta.nlink() as f64));
    }
    #[cfg(not(unix))]
    {
        o.properties.insert("link-count".into(), Value::F64(1.0));
    }
    o.properties
        .insert("size".into(), Value::F64(meta.len() as f64));
    if let Ok(t) = meta.modified() {
        o.properties.insert(
            "data-modification-timestamp".into(),
            Value::F64(ms_since_epoch(t)),
        );
    }
    if let Ok(t) = meta.accessed() {
        o.properties.insert(
            "data-access-timestamp".into(),
            Value::F64(ms_since_epoch(t)),
        );
    }
    if let Ok(t) = meta.created() {
        o.properties.insert(
            "status-change-timestamp".into(),
            Value::F64(ms_since_epoch(t)),
        );
    }
    Value::Object(vybe_bytecode::heap::alloc(o))
}

// ── path resolution ───────────────────────────────────────────────
//
// WASI paths are interpreted relative to the parent descriptor.
// `path-flags::symlink-follow` (bit 0) controls whether symlinks
// in the resolved path are followed; the std::fs APIs follow by
// default, so the flag is currently informational.

fn resolve_child_path(parent_id: u32, child: &str) -> Result<PathBuf, &'static str> {
    let registry = registry().lock().unwrap();
    let parent = registry
        .descriptors
        .get(&parent_id)
        .ok_or("bad-descriptor")?;
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

// Decode a WASI `new-timestamp` variant into an optional SystemTime.
// Null → no-change (None); Object{seconds,nanoseconds} → that instant; fallback → now.
fn new_timestamp(v: Option<&Value>) -> Option<SystemTime> {
    match v {
        None | Some(Value::Null) => None,
        Some(Value::Object(obj)) => {
            let inner = obj.lock().unwrap();
            let secs = inner
                .properties
                .get("seconds")
                .map(|v| v.as_f64() as u64)
                .unwrap_or(0);
            let ns = inner
                .properties
                .get("nanoseconds")
                .map(|v| v.as_f64() as u32)
                .unwrap_or(0);
            Some(SystemTime::UNIX_EPOCH + Duration::new(secs, ns))
        }
        _ => Some(SystemTime::now()),
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
    vm.register_host_fn(
        "wasi:filesystem/preopens",
        "get-directories",
        Box::new(|_ctx, _args| {
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
            let pair = Value::Object(vybe_bytecode::heap::alloc(Object::new_array(pair_elements)));
            Value::Object(vybe_bytecode::heap::alloc(Object::new_array(vec![pair])))
        }),
    );
}

fn register_types(vm: &mut VM) {
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.open-at",
        Box::new(|_ctx, args| {
            let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let _path_flags = i32_arg(args, 1, 0);
            let Some(child) = s_arg(args, 2) else {
                return err("invalid");
            };
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
            if exists && create && exclusive {
                return err("exist");
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
                if exclusive {
                    opts.create_new(true);
                }
                if truncate {
                    opts.truncate(true);
                }
                if let Err(e) = opts.open(&resolved) {
                    return err(map_io_error(&e));
                }
            } else if exists && truncate && resolved.is_file() {
                if let Err(e) = OpenOptions::new()
                    .write(true)
                    .truncate(true)
                    .open(&resolved)
                {
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
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.stat",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
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
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.stat-at",
        Box::new(|_ctx, args| {
            let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let _path_flags = i32_arg(args, 1, 0);
            let Some(child) = s_arg(args, 2) else {
                return err("invalid");
            };
            let resolved = match resolve_child_path(parent_id, &child) {
                Ok(p) => p,
                Err(code) => return err(code),
            };
            match std::fs::metadata(&resolved) {
                Ok(meta) => build_stat(&meta),
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.get-type",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
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
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.read-directory",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
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
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]directory-entry-stream.read-directory-entry",
        Box::new(|_ctx, args| {
            let Some(stream_id) = resource_id(&args[0].clone(), KIND_DIR_STREAM) else {
                return err("bad-descriptor");
            };
            let mut reg = registry().lock().unwrap();
            let Some(stream) = reg.dir_streams.get_mut(&stream_id) else {
                return err("bad-descriptor");
            };
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
                    o.properties
                        .insert("type".into(), Value::String(Arc::from(entry_type)));
                    o.properties
                        .insert("name".into(), Value::String(Arc::from(name.as_str())));
                    Value::Object(vybe_bytecode::heap::alloc(o))
                }
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.create-directory-at",
        Box::new(|_ctx, args| {
            let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(child) = s_arg(args, 1) else {
                return err("invalid");
            };
            let resolved = match resolve_child_path(parent_id, &child) {
                Ok(p) => p,
                Err(code) => return err(code),
            };
            match std::fs::create_dir(&resolved) {
                Ok(_) => Value::Null,
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.unlink-file-at",
        Box::new(|_ctx, args| {
            let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(child) = s_arg(args, 1) else {
                return err("invalid");
            };
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
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.remove-directory-at",
        Box::new(|_ctx, args| {
            let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(child) = s_arg(args, 1) else {
                return err("invalid");
            };
            let resolved = match resolve_child_path(parent_id, &child) {
                Ok(p) => p,
                Err(code) => return err(code),
            };
            match std::fs::remove_dir(&resolved) {
                Ok(_) => Value::Null,
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.rename-at",
        Box::new(|_ctx, args| {
            let Some(old_parent) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(old_child) = s_arg(args, 1) else {
                return err("invalid");
            };
            let Some(new_parent) = resource_id(&args[2].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(new_child) = s_arg(args, 3) else {
                return err("invalid");
            };
            let old_path = match resolve_child_path(old_parent, &old_child) {
                Ok(p) => p,
                Err(c) => return err(c),
            };
            let new_path = match resolve_child_path(new_parent, &new_child) {
                Ok(p) => p,
                Err(c) => return err(c),
            };
            match std::fs::rename(&old_path, &new_path) {
                Ok(_) => Value::Null,
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.readlink-at",
        Box::new(|_ctx, args| {
            let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(child) = s_arg(args, 1) else {
                return err("invalid");
            };
            let resolved = match resolve_child_path(parent_id, &child) {
                Ok(p) => p,
                Err(c) => return err(c),
            };
            match std::fs::read_link(&resolved) {
                Ok(target) => Value::String(Arc::from(target.to_string_lossy().as_ref())),
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.is-same-object",
        Box::new(|_ctx, args| {
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
        }),
    );

    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.read-via-stream",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
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
                reg.input_streams.insert(
                    stream_id,
                    InputStreamKind::File {
                        file,
                        position: offset,
                    },
                );
                stream_id
            };
            make_resource(KIND_INPUT_STREAM, stream_id)
        }),
    );

    // write-via-stream(offset) → output-stream
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.write-via-stream",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let offset = u64_arg(args, 1, 0);
            let path = {
                let reg = registry().lock().unwrap();
                match reg.descriptors.get(&id) {
                    Some(DescriptorKind::File(p)) => p.clone(),
                    Some(DescriptorKind::Directory(_)) => return err("is-directory"),
                    None => return err("bad-descriptor"),
                }
            };
            let mut file = match OpenOptions::new().write(true).open(&path) {
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
                let sid = reg.alloc_id();
                reg.output_streams
                    .insert(sid, OutputStreamKind::File { file });
                sid
            };
            make_resource(KIND_OUTPUT_STREAM, stream_id)
        }),
    );

    // append-via-stream() → output-stream
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.append-via-stream",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let path = {
                let reg = registry().lock().unwrap();
                match reg.descriptors.get(&id) {
                    Some(DescriptorKind::File(p)) => p.clone(),
                    Some(DescriptorKind::Directory(_)) => return err("is-directory"),
                    None => return err("bad-descriptor"),
                }
            };
            let stream_id = {
                let mut reg = registry().lock().unwrap();
                let sid = reg.alloc_id();
                reg.output_streams
                    .insert(sid, OutputStreamKind::Append(path));
                sid
            };
            make_resource(KIND_OUTPUT_STREAM, stream_id)
        }),
    );

    // advise(offset, length, advice) → result — advisory access hint, best-effort stub
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.advise",
        Box::new(|_ctx, args| {
            if resource_id(&args[0].clone(), KIND_DESCRIPTOR).is_none() {
                return err("bad-descriptor");
            }
            Value::Null
        }),
    );

    // sync-data() → result
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.sync-data",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let path = {
                let reg = registry().lock().unwrap();
                match reg.descriptors.get(&id) {
                    Some(DescriptorKind::File(p)) => p.clone(),
                    Some(DescriptorKind::Directory(_)) => return Value::Null,
                    None => return err("bad-descriptor"),
                }
            };
            match File::open(&path).and_then(|f| f.sync_data()) {
                Ok(_) => Value::Null,
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    // get-flags() → result<descriptor-flags, error-code>
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.get-flags",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let is_dir = {
                let reg = registry().lock().unwrap();
                matches!(reg.descriptors.get(&id), Some(DescriptorKind::Directory(_)))
            };
            let mut flags = Object::new();
            flags.properties.insert("read".into(), Value::Bool(true));
            flags
                .properties
                .insert("write".into(), Value::Bool(!is_dir));
            flags
                .properties
                .insert("file-integrity-sync".into(), Value::Bool(false));
            flags
                .properties
                .insert("data-integrity-sync".into(), Value::Bool(false));
            flags
                .properties
                .insert("requested-write-sync".into(), Value::Bool(false));
            flags
                .properties
                .insert("mutate-directory".into(), Value::Bool(false));
            Value::Object(vybe_bytecode::heap::alloc(flags))
        }),
    );

    // set-size(size) → result
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.set-size",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let size = u64_arg(args, 1, 0);
            let path = {
                let reg = registry().lock().unwrap();
                match reg.descriptors.get(&id) {
                    Some(DescriptorKind::File(p)) => p.clone(),
                    Some(DescriptorKind::Directory(_)) => return err("is-directory"),
                    None => return err("bad-descriptor"),
                }
            };
            match OpenOptions::new()
                .write(true)
                .open(&path)
                .and_then(|f| f.set_len(size))
            {
                Ok(_) => Value::Null,
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    // set-times(data-access-timestamp, data-modification-timestamp) → result
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.set-times",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let atime = new_timestamp(args.get(1));
            let mtime = new_timestamp(args.get(2));
            let path = {
                let reg = registry().lock().unwrap();
                match reg.descriptors.get(&id) {
                    Some(DescriptorKind::File(p)) => p.clone(),
                    Some(DescriptorKind::Directory(_)) => return Value::Null, // dirs: no-op
                    None => return err("bad-descriptor"),
                }
            };
            match OpenOptions::new().write(true).open(&path) {
                Ok(file) => {
                    let mut times = FileTimes::new();
                    if let Some(t) = atime {
                        times = times.set_accessed(t);
                    }
                    if let Some(t) = mtime {
                        times = times.set_modified(t);
                    }
                    match file.set_times(times) {
                        Ok(_) => Value::Null,
                        Err(e) => err(map_io_error(&e)),
                    }
                }
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    // read(length, offset) → result<tuple<list<u8>, bool>, error-code>  (pread)
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.read",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let length = u64_arg(args, 1, 0);
            let offset = u64_arg(args, 2, 0);
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
            if let Err(e) = file.seek(SeekFrom::Start(offset)) {
                return err(map_io_error(&e));
            }
            let cap = length.min(64 * 1024) as usize;
            let mut buf = vec![0u8; cap];
            match file.read(&mut buf) {
                Ok(n) => {
                    buf.truncate(n);
                    let eof = n == 0 || n < cap;
                    let bytes: Vec<Value> = buf.into_iter().map(|b| Value::I32(b as i32)).collect();
                    let bytes_val = Value::Object(vybe_bytecode::heap::alloc(Object::new_array(bytes)));
                    let tuple = Object::new_array(vec![bytes_val, Value::Bool(eof)]);
                    Value::Object(vybe_bytecode::heap::alloc(tuple))
                }
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    // write(bytes, offset) → result<u64, error-code>  (pwrite)
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.write",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let bytes_val = args.get(1).cloned().unwrap_or(Value::Null);
            let offset = u64_arg(args, 2, 0);
            let bytes: Vec<u8> = if let Value::Object(arr) = &bytes_val {
                let inner = arr.lock().unwrap();
                if let vybe_bytecode::value::ObjectKind::Array(ref elems) = inner.kind {
                    elems.iter().map(|v| v.as_f64() as u8).collect()
                } else {
                    return err("invalid");
                }
            } else {
                return err("invalid");
            };
            let path = {
                let reg = registry().lock().unwrap();
                match reg.descriptors.get(&id) {
                    Some(DescriptorKind::File(p)) => p.clone(),
                    Some(DescriptorKind::Directory(_)) => return err("is-directory"),
                    None => return err("bad-descriptor"),
                }
            };
            let mut file = match OpenOptions::new().write(true).open(&path) {
                Ok(f) => f,
                Err(e) => return err(map_io_error(&e)),
            };
            if let Err(e) = file.seek(SeekFrom::Start(offset)) {
                return err(map_io_error(&e));
            }
            match file.write_all(&bytes) {
                Ok(_) => Value::F64(bytes.len() as f64),
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    // sync() → result
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.sync",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let path = {
                let reg = registry().lock().unwrap();
                match reg.descriptors.get(&id) {
                    Some(DescriptorKind::File(p)) => p.clone(),
                    Some(DescriptorKind::Directory(_)) => return Value::Null,
                    None => return err("bad-descriptor"),
                }
            };
            match File::open(&path).and_then(|f| f.sync_all()) {
                Ok(_) => Value::Null,
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    // set-times-at(path-flags, path, atime, mtime) → result
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.set-times-at",
        Box::new(|_ctx, args| {
            let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(child) = s_arg(args, 2) else {
                return err("invalid");
            };
            let atime = new_timestamp(args.get(3));
            let mtime = new_timestamp(args.get(4));
            let resolved = match resolve_child_path(parent_id, &child) {
                Ok(p) => p,
                Err(c) => return err(c),
            };
            match OpenOptions::new().write(true).open(&resolved) {
                Ok(file) => {
                    let mut times = FileTimes::new();
                    if let Some(t) = atime {
                        times = times.set_accessed(t);
                    }
                    if let Some(t) = mtime {
                        times = times.set_modified(t);
                    }
                    match file.set_times(times) {
                        Ok(_) => Value::Null,
                        Err(e) => err(map_io_error(&e)),
                    }
                }
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    // link-at(old-path-flags, old-path, new-descriptor, new-path) → result
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.link-at",
        Box::new(|_ctx, args| {
            let Some(old_parent) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(old_child) = s_arg(args, 2) else {
                return err("invalid");
            };
            let Some(new_parent) = resource_id(&args[3].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(new_child) = s_arg(args, 4) else {
                return err("invalid");
            };
            let old_path = match resolve_child_path(old_parent, &old_child) {
                Ok(p) => p,
                Err(c) => return err(c),
            };
            let new_path = match resolve_child_path(new_parent, &new_child) {
                Ok(p) => p,
                Err(c) => return err(c),
            };
            #[cfg(unix)]
            {
                match std::fs::hard_link(&old_path, &new_path) {
                    Ok(_) => Value::Null,
                    Err(e) => err(map_io_error(&e)),
                }
            }
            #[cfg(not(unix))]
            {
                match std::fs::hard_link(&old_path, &new_path) {
                    Ok(_) => Value::Null,
                    Err(e) => err(map_io_error(&e)),
                }
            }
        }),
    );

    // symlink-at(old-path, new-descriptor, new-path) → result
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.symlink-at",
        Box::new(|_ctx, args| {
            let Some(parent) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(old_path_str) = s_arg(args, 1) else {
                return err("invalid");
            };
            let Some(new_parent) = resource_id(&args[2].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(new_child) = s_arg(args, 3) else {
                return err("invalid");
            };
            let _old_path = match resolve_child_path(parent, &old_path_str) {
                Ok(p) => p,
                Err(c) => return err(c),
            };
            let new_path = match resolve_child_path(new_parent, &new_child) {
                Ok(p) => p,
                Err(c) => return err(c),
            };
            #[cfg(unix)]
            {
                match std::os::unix::fs::symlink(&old_path_str, &new_path) {
                    Ok(_) => Value::Null,
                    Err(e) => err(map_io_error(&e)),
                }
            }
            #[cfg(windows)]
            {
                // On Windows we'd need to know if target is file or dir; stub as error.
                err("unsupported")
            }
            #[cfg(not(any(unix, windows)))]
            {
                err("unsupported")
            }
        }),
    );

    // metadata-hash() → result<metadata-hash-value, error-code>
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.metadata-hash",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let path = {
                let reg = registry().lock().unwrap();
                match reg.descriptors.get(&id) {
                    Some(DescriptorKind::File(p)) | Some(DescriptorKind::Directory(p)) => p.clone(),
                    None => return err("bad-descriptor"),
                }
            };
            match std::fs::metadata(&path) {
                Ok(meta) => {
                    let hash = meta.len().wrapping_mul(0x9e37_79b9)
                        ^ ms_since_epoch(meta.modified().unwrap_or(SystemTime::UNIX_EPOCH)) as u64;
                    let mut o = Object::new();
                    o.properties
                        .insert("lower".into(), Value::F64((hash & 0xffff_ffff) as f64));
                    o.properties
                        .insert("upper".into(), Value::F64((hash >> 32) as f64));
                    Value::Object(vybe_bytecode::heap::alloc(o))
                }
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    // metadata-hash-at(path-flags, path) → result<metadata-hash-value, error-code>
    vm.register_host_fn(
        "wasi:filesystem/types",
        "[method]descriptor.metadata-hash-at",
        Box::new(|_ctx, args| {
            let Some(parent_id) = resource_id(&args[0].clone(), KIND_DESCRIPTOR) else {
                return err("bad-descriptor");
            };
            let Some(child) = s_arg(args, 2) else {
                return err("invalid");
            };
            let resolved = match resolve_child_path(parent_id, &child) {
                Ok(p) => p,
                Err(c) => return err(c),
            };
            match std::fs::metadata(&resolved) {
                Ok(meta) => {
                    let hash = meta.len().wrapping_mul(0x9e37_79b9)
                        ^ ms_since_epoch(meta.modified().unwrap_or(SystemTime::UNIX_EPOCH)) as u64;
                    let mut o = Object::new();
                    o.properties
                        .insert("lower".into(), Value::F64((hash & 0xffff_ffff) as f64));
                    o.properties
                        .insert("upper".into(), Value::F64((hash >> 32) as f64));
                    Value::Object(vybe_bytecode::heap::alloc(o))
                }
                Err(e) => err(map_io_error(&e)),
            }
        }),
    );

    // filesystem-error-code(err) → option<error-code>
    // Converts a wasi:io/error into a filesystem error-code if it came from this module.
    vm.register_host_fn(
        "wasi:filesystem/types",
        "filesystem-error-code",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(obj)) = args.first() {
                let inner = obj.lock().unwrap();
                if let Some(code) = inner.properties.get("__wasi_error") {
                    return code.clone();
                }
            }
            Value::Null
        }),
    );
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
    vm.register_host_fn(
        "wasi:io/streams",
        "[method]input-stream.blocking-read",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_INPUT_STREAM) else {
                return err("bad-descriptor");
            };
            let max = u64_arg(args, 1, u64::MAX);
            let mut reg = registry().lock().unwrap();
            let Some(stream) = reg.input_streams.get_mut(&id) else {
                return err("bad-descriptor");
            };
            match stream {
                InputStreamKind::File { file, position } => {
                    let cap = max.min(64 * 1024) as usize;
                    let mut buf = vec![0u8; cap];
                    match file.read(&mut buf) {
                        Ok(n) => {
                            buf.truncate(n);
                            *position += n as u64;
                            let elements: Vec<Value> =
                                buf.into_iter().map(|b| Value::I32(b as i32)).collect();
                            Value::Object(vybe_bytecode::heap::alloc(Object::new_array(elements)))
                        }
                        Err(e) => err(map_io_error(&e)),
                    }
                }
                InputStreamKind::Buffer { data, position } => {
                    let remaining = data.len().saturating_sub(*position);
                    let cap = (max as usize).min(remaining).min(64 * 1024);
                    let slice = &data[*position..*position + cap];
                    let elements: Vec<Value> =
                        slice.iter().map(|b| Value::I32(*b as i32)).collect();
                    *position += cap;
                    Value::Object(vybe_bytecode::heap::alloc(Object::new_array(elements)))
                }
            }
        }),
    );

    vm.register_host_fn(
        "wasi:io/streams",
        "[method]input-stream.read",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_INPUT_STREAM) else {
                return err("bad-descriptor");
            };
            let max = u64_arg(args, 1, u64::MAX);
            let mut reg = registry().lock().unwrap();
            let Some(stream) = reg.input_streams.get_mut(&id) else {
                return err("bad-descriptor");
            };
            match stream {
                InputStreamKind::File { file, position } => {
                    let cap = max.min(64 * 1024) as usize;
                    let mut buf = vec![0u8; cap];
                    match file.read(&mut buf) {
                        Ok(n) => {
                            buf.truncate(n);
                            *position += n as u64;
                            let elements: Vec<Value> =
                                buf.into_iter().map(|b| Value::I32(b as i32)).collect();
                            Value::Object(vybe_bytecode::heap::alloc(Object::new_array(elements)))
                        }
                        Err(e) => err(map_io_error(&e)),
                    }
                }
                InputStreamKind::Buffer { data, position } => {
                    let remaining = data.len().saturating_sub(*position);
                    let cap = (max as usize).min(remaining).min(64 * 1024);
                    let slice = &data[*position..*position + cap];
                    let elements: Vec<Value> =
                        slice.iter().map(|b| Value::I32(*b as i32)).collect();
                    *position += cap;
                    Value::Object(vybe_bytecode::heap::alloc(Object::new_array(elements)))
                }
            }
        }),
    );

    // output-stream write/flush/check-write/subscribe
    vm.register_host_fn(
        "wasi:io/streams",
        "[method]output-stream.write",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_OUTPUT_STREAM) else {
                return err("bad-descriptor");
            };
            let bytes_val = args.get(1).cloned().unwrap_or(Value::Null);
            let bytes: Vec<u8> = if let Value::Object(arr) = &bytes_val {
                let inner = arr.lock().unwrap();
                if let vybe_bytecode::value::ObjectKind::Array(ref elems) = inner.kind {
                    elems.iter().map(|v| v.as_f64() as u8).collect()
                } else {
                    return err("invalid");
                }
            } else {
                return err("invalid");
            };
            let mut reg = registry().lock().unwrap();
            match reg.output_streams.get_mut(&id) {
                Some(OutputStreamKind::File { file }) => match file.write_all(&bytes) {
                    Ok(_) => Value::F64(bytes.len() as f64),
                    Err(e) => err(map_io_error(&e)),
                },
                Some(OutputStreamKind::Append(path)) => {
                    let path = path.clone();
                    match OpenOptions::new()
                        .append(true)
                        .open(&path)
                        .and_then(|mut f| f.write_all(&bytes))
                    {
                        Ok(_) => Value::F64(bytes.len() as f64),
                        Err(e) => err(map_io_error(&e)),
                    }
                }
                None => err("bad-descriptor"),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:io/streams",
        "[method]output-stream.blocking-write-and-flush",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_OUTPUT_STREAM) else {
                return err("bad-descriptor");
            };
            let bytes_val = args.get(1).cloned().unwrap_or(Value::Null);
            let bytes: Vec<u8> = if let Value::Object(arr) = &bytes_val {
                let inner = arr.lock().unwrap();
                if let vybe_bytecode::value::ObjectKind::Array(ref elems) = inner.kind {
                    elems.iter().map(|v| v.as_f64() as u8).collect()
                } else {
                    return err("invalid");
                }
            } else {
                return err("invalid");
            };
            let mut reg = registry().lock().unwrap();
            match reg.output_streams.get_mut(&id) {
                Some(OutputStreamKind::File { file }) => {
                    if file.write_all(&bytes).is_err() || file.flush().is_err() {
                        return err("io");
                    }
                    Value::Null
                }
                Some(OutputStreamKind::Append(path)) => {
                    let path = path.clone();
                    match OpenOptions::new()
                        .append(true)
                        .open(&path)
                        .and_then(|mut f| {
                            f.write_all(&bytes)?;
                            f.flush()
                        }) {
                        Ok(_) => Value::Null,
                        Err(e) => err(map_io_error(&e)),
                    }
                }
                None => err("bad-descriptor"),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:io/streams",
        "[method]output-stream.check-write",
        Box::new(|_ctx, args| {
            if resource_id(&args[0].clone(), KIND_OUTPUT_STREAM).is_none() {
                return err("bad-descriptor");
            }
            Value::F64(65536.0) // always-ready: 64 KiB budget
        }),
    );

    vm.register_host_fn(
        "wasi:io/streams",
        "[method]output-stream.flush",
        Box::new(|_ctx, args| {
            let Some(id) = resource_id(&args[0].clone(), KIND_OUTPUT_STREAM) else {
                return err("bad-descriptor");
            };
            let mut reg = registry().lock().unwrap();
            match reg.output_streams.get_mut(&id) {
                Some(OutputStreamKind::File { file }) => {
                    let _ = file.flush();
                    Value::Null
                }
                Some(OutputStreamKind::Append(_)) => Value::Null,
                None => err("bad-descriptor"),
            }
        }),
    );

    vm.register_host_fn(
        "wasi:io/streams",
        "[method]output-stream.subscribe",
        Box::new(|_ctx, args| {
            if resource_id(&args[0].clone(), KIND_OUTPUT_STREAM).is_none() {
                return err("bad-descriptor");
            }
            // Output streams are always ready in our sync model — return a pre-resolved pollable.
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("Pollable")));
            obj.properties.insert("__ready".into(), Value::Bool(true));
            Value::Object(vybe_bytecode::heap::alloc(obj))
        }),
    );
}

// ── Test-only helpers ─────────────────────────────────────────────
//
// Production code receives preopened directories via
// `get-directories`. Tests need to build a descriptor for a known
// scratch path; this opens that backdoor under a `__test_*` name so
// there's no chance of guest code stumbling onto it accidentally
// (none of the language profiles route imports to `__test_*`).

fn register_test_helpers(vm: &mut VM) {
    vm.register_host_fn(
        "wasi:filesystem/types",
        "__test_open_root",
        Box::new(|_ctx, args| {
            let Some(path_str) = s_arg(args, 0) else {
                return err("invalid");
            };
            let path = PathBuf::from(&path_str);
            if !path.is_dir() {
                return err("not-directory");
            }
            let id = {
                let mut reg = registry().lock().unwrap();
                let id = reg.alloc_id();
                reg.descriptors.insert(id, DescriptorKind::Directory(path));
                id
            };
            make_resource(KIND_DESCRIPTOR, id)
        }),
    );
}

// Suppress unused-constant warnings for the descriptor-flag bits.
#[allow(dead_code)]
const _: i32 = DESC_READ + DESC_WRITE;
