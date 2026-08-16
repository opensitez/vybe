//! `wasi:io` — unified `[method]resource.fn` forms for WASI 0.2.12 backward compat.
//!
//! Covers:
//!   - `wasi:io/streams`  — input-stream and output-stream resources
//!   - `wasi:io/poll`     — pollable resource + poll() free function
//!   - `wasi:io/error`    — error resource
//!
//! Each handler dispatches on the resource shape:
//!   • File streams:   `__wasi_kind == "input-stream"/"output-stream"` + `__wasi_id`
//!   • Socket streams: `__socket_id` present (set by wasi:sockets/tcp.finish-connect)
//!   • fd streams:     `fd == 0/1/2` (stdin/stdout/stderr from wasi:cli/std*)
//!   • Pollables:      `__type == "Pollable"`
//!   • Errors:         `__wasi_error` property

use std::io::{Read, Write};
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{Duration, Instant};

use vybe_runtime::value::Object;
use vybe_runtime::vm::{HostFnDecl, ResourceBinding, ResourceMemberKind};
use vybe_runtime::{FuncSig, HostContext, VM, ValType, Value};

use super::filesystem as fs;
use crate::sockets as skt;

/// The four resources `wasi:io` defines. These names are already in every
/// registration — `"[method]input-stream.read"` is WIT's own resource-method
/// spelling — but as part of a STRING. Naming them here lets the declaration
/// carry the same fact structurally, so `own` vs `borrow` and "this is a
/// destructor" become checkable rather than a convention in a name.
const INPUT_STREAM: &str = "input-stream";
const OUTPUT_STREAM: &str = "output-stream";
const POLLABLE: &str = "pollable";
const ERROR: &str = "error";

fn borrowed(resource: &str) -> ValType {
    ValType::Borrow(resource.to_string())
}

/// A WASI `result<T, stream-error>`. The error arm is `Any`: `stream-error` is
/// a variant carrying an `error` RESOURCE in one arm, and `ValType` has no
/// variant. Declaring `Result` at all is still worth it — it says the call can
/// fail, which is the half every caller has to handle.
fn wasi_result(ok: ValType) -> ValType {
    ValType::Result(Box::new(ok), Box::new(ValType::Any))
}

/// Register a `wasi:*` function WITH its WIT signature.
///
/// The WIT name (`"[method]output-stream.splice"`) stays the registry key,
/// because that is what emitters import. The Component Model name is the bare
/// member (`"splice"`), which is what it is called ON the resource.
///
/// **The signature comes from the WIT, not from the closure.** Several closures
/// here ignore their `self` handle — `flush`, `pollable.ready` and every
/// `[resource-drop]` read no arguments at all, because this host has one
/// instance and does not need the handle to find it. The call sites still pass
/// it (`pollable.block` is emitted with argc 1, `blocking-read` with 2), so the
/// contract is WIT's and the closure merely does not consult it. That gap is
/// exactly what a declaration is for.
fn wasi_fn(
    vm: &mut VM,
    module: &str,
    name: &str,
    resource: Option<(&str, ResourceMemberKind)>,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    let member = name
        .rsplit('.')
        .next()
        .unwrap_or(name)
        .trim_start_matches("[resource-drop]")
        .to_string();
    let mut decl = HostFnDecl::new(module, name, call).with_sig(FuncSig {
        name: member,
        params,
        results,
    });
    if let Some((res, kind)) = resource {
        decl = decl.resource_member(ResourceBinding {
            resource: res.to_string(),
            kind,
            // A destructor CONSUMES the handle — `own<T>`, not `borrow<T>`.
            // Every other member borrows: reading a stream does not end it.
            borrows_self: !matches!(kind, ResourceMemberKind::Destructor),
        });
    }
    vm.register_host(decl);
}

pub fn register(vm: &mut VM) {
    register_input_stream(vm);
    register_output_stream(vm);
    register_pollable(vm);
    register_error(vm);
}

// ── helpers ───────────────────────────────────────────────────────────────────

fn as_obj(v: &Value) -> Option<Arc<Mutex<Object>>> {
    if let Value::Object(o) = v {
        Some(o.clone())
    } else {
        None
    }
}

fn u64_arg(args: &[Value], idx: usize) -> u64 {
    match args.get(idx) {
        Some(Value::F64(n)) => *n as u64,
        Some(Value::I32(n)) => *n as u64,
        Some(Value::I64(n)) => *n as u64,
        _ => 0,
    }
}

fn stream_fd(v: &Value) -> Option<i32> {
    if let Value::Object(obj) = v {
        if let Some(Value::I32(fd)) = obj.lock().unwrap().properties.get("fd") {
            return Some(*fd);
        }
    }
    None
}

fn read_file_stream(id: u32, max: usize) -> Value {
    let mut reg = fs::registry().lock().unwrap();
    let Some(stream) = reg.input_streams.get_mut(&id) else {
        return fs::err("bad-descriptor");
    };
    match stream {
        fs::InputStreamKind::File { file, position } => {
            let cap = max.min(64 * 1024);
            let mut buf = vec![0u8; cap];
            match file.read(&mut buf) {
                Ok(n) => {
                    buf.truncate(n);
                    *position += n as u64;
                    let elems: Vec<Value> = buf.into_iter().map(|b| Value::I32(b as i32)).collect();
                    Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))
                }
                Err(e) => fs::err(fs::map_io_error(&e)),
            }
        }
        fs::InputStreamKind::Buffer { data, position } => {
            let remaining = data.len().saturating_sub(*position);
            let cap = max.min(remaining).min(64 * 1024);
            let elems: Vec<Value> = data[*position..*position + cap]
                .iter()
                .map(|b| Value::I32(*b as i32))
                .collect();
            *position += cap;
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(elems)))
        }
    }
}

fn write_file_stream(id: u32, bytes: &[u8], flush: bool) -> Value {
    use std::fs::OpenOptions;
    let mut reg = fs::registry().lock().unwrap();
    match reg.output_streams.get_mut(&id) {
        Some(fs::OutputStreamKind::File { file }) => {
            if file.write_all(bytes).is_err() {
                return fs::err("io");
            }
            if flush {
                let _ = file.flush();
            }
            Value::F64(bytes.len() as f64)
        }
        Some(fs::OutputStreamKind::Append(path)) => {
            let path = path.clone();
            match OpenOptions::new()
                .append(true)
                .open(&path)
                .and_then(|mut f| f.write_all(bytes))
            {
                Ok(_) => Value::F64(bytes.len() as f64),
                Err(e) => fs::err(fs::map_io_error(&e)),
            }
        }
        None => fs::err("bad-descriptor"),
    }
}

fn flush_file_stream(id: u32) -> Value {
    let mut reg = fs::registry().lock().unwrap();
    match reg.output_streams.get_mut(&id) {
        Some(fs::OutputStreamKind::File { file }) => {
            let _ = file.flush();
            Value::Null
        }
        Some(fs::OutputStreamKind::Append(_)) => Value::Null,
        None => fs::err("bad-descriptor"),
    }
}

fn make_ready_pollable() -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Pollable")));
    obj.properties.insert("__ready".into(), Value::Bool(true));
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn read_bytes_from_stream(stream: &Value, len: usize, blocking: bool) -> Vec<u8> {
    if let Some(id) = fs::resource_id(stream, fs::KIND_INPUT_STREAM) {
        return match read_file_stream(id, len) {
            Value::Object(arr) => skt::bytes_from_value(&Value::Object(arr)),
            _ => vec![],
        };
    }
    if let Some(obj) = as_obj(stream) {
        if skt::stream_socket_id(&obj).is_some() {
            return skt::bytes_from_value(&skt::read_stream_bytes(&obj, len, blocking));
        }
    }
    if let Some(fd) = stream_fd(stream) {
        if fd == 0 {
            if !blocking {
                return vec![];
            }
            let mut buf = vec![0u8; len.min(4096)];
            return match std::io::stdin().read(&mut buf) {
                Ok(n) => {
                    buf.truncate(n);
                    buf
                }
                Err(_) => vec![],
            };
        }
    }
    vec![]
}

fn write_bytes_to_stream(stream: &Value, bytes: &[u8], flush: bool) -> bool {
    if let Some(id) = fs::resource_id(stream, fs::KIND_OUTPUT_STREAM) {
        return !is_err_value(&write_file_stream(id, bytes, flush));
    }
    if let Some(obj) = as_obj(stream) {
        if skt::stream_socket_id(&obj).is_some() {
            skt::write_stream_bytes(&obj, bytes, flush);
            return true;
        }
    }
    if let Some(fd) = stream_fd(stream) {
        let result = match fd {
            1 => std::io::stdout().write_all(bytes),
            2 => std::io::stderr().write_all(bytes),
            _ => return false,
        };
        if flush {
            let _ = match fd {
                1 => std::io::stdout().flush(),
                2 => std::io::stderr().flush(),
                _ => Ok(()),
            };
        }
        return result.is_ok();
    }
    false
}

fn is_err_value(v: &Value) -> bool {
    if let Value::Object(obj) = v {
        obj.lock().unwrap().properties.contains_key("__wasi_error")
    } else {
        false
    }
}

fn collect_ready_indices(pollables: &[Value]) -> Vec<usize> {
    pollables
        .iter()
        .enumerate()
        .filter_map(|(i, v)| {
            let obj = as_obj(v)?;
            let ready = {
                let locked = obj.lock().unwrap();
                if let Some(Value::Bool(r)) = locked.properties.get("__ready") {
                    *r
                } else {
                    drop(locked);
                    skt::pollable_ready(&obj)
                }
            };
            ready.then_some(i)
        })
        .collect()
}

// ── wasi:io/streams — input-stream ───────────────────────────────────────────

fn register_input_stream(vm: &mut VM) {
    // [method]input-stream.read(self, len: u64) -> result<list<u8>, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]input-stream.read",
        Some((INPUT_STREAM, ResourceMemberKind::Method)),
        // `list<u8>` is `List(I32)`: `ValType` has no `u8`, and widening the
        // ELEMENT is honest where inventing a narrower type would not be.
        vec![borrowed(INPUT_STREAM), ValType::I64],
        vec![wasi_result(ValType::List(Box::new(ValType::I32)))],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            let max = u64_arg(args, 1) as usize;
            if let Some(id) = fs::resource_id(&stream, fs::KIND_INPUT_STREAM) {
                return read_file_stream(id, max);
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::read_stream_bytes(&obj, max, false);
                }
            }
            if stream_fd(&stream) == Some(0) {
                return skt::value_array(vec![]);
            }
            fs::err("bad-descriptor")
        }),
    );

    // [method]input-stream.blocking-read(self, len: u64) -> result<list<u8>, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]input-stream.blocking-read",
        Some((INPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(INPUT_STREAM), ValType::I64],
        vec![wasi_result(ValType::List(Box::new(ValType::I32)))],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            let max = (u64_arg(args, 1) as usize).max(1);
            if let Some(id) = fs::resource_id(&stream, fs::KIND_INPUT_STREAM) {
                return read_file_stream(id, max);
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::read_stream_bytes(&obj, max, true);
                }
            }
            if stream_fd(&stream) == Some(0) {
                use std::io::BufRead;
                let mut line = String::new();
                let stdin = std::io::stdin();
                let _ = stdin.lock().read_line(&mut line);
                let trimmed = line.trim_end_matches('\n').trim_end_matches('\r');
                return Value::String(Arc::from(trimmed));
            }
            fs::err("bad-descriptor")
        }),
    );

    // [method]input-stream.skip(self, num: u64) -> result<u64, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]input-stream.skip",
        Some((INPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(INPUT_STREAM), ValType::I64],
        vec![wasi_result(ValType::I64)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            let len = u64_arg(args, 1) as usize;
            if let Some(id) = fs::resource_id(&stream, fs::KIND_INPUT_STREAM) {
                return match read_file_stream(id, len) {
                    Value::Object(arr) => Value::I64(skt::array_len(&arr) as i64),
                    other => other,
                };
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return match skt::read_stream_bytes(&obj, len, false) {
                        Value::Object(arr) => Value::I64(skt::array_len(&arr) as i64),
                        other => other,
                    };
                }
            }
            Value::I64(0)
        }),
    );

    // [method]input-stream.blocking-skip(self, num: u64) -> result<u64, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]input-stream.blocking-skip",
        Some((INPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(INPUT_STREAM), ValType::I64],
        vec![wasi_result(ValType::I64)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            let len = u64_arg(args, 1) as usize;
            if let Some(id) = fs::resource_id(&stream, fs::KIND_INPUT_STREAM) {
                return match read_file_stream(id, len) {
                    Value::Object(arr) => Value::I64(skt::array_len(&arr) as i64),
                    other => other,
                };
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return match skt::read_stream_bytes(&obj, len, true) {
                        Value::Object(arr) => Value::I64(skt::array_len(&arr) as i64),
                        other => other,
                    };
                }
            }
            Value::I64(0)
        }),
    );

    // [method]input-stream.subscribe(self) -> pollable
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]input-stream.subscribe",
        Some((INPUT_STREAM, ResourceMemberKind::Method)),
        // Returns an OWNED pollable: the caller gets a new handle it must drop.
        // The stream itself is only borrowed — subscribing does not consume it.
        vec![borrowed(INPUT_STREAM)],
        vec![ValType::Own(POLLABLE.to_string())],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            if fs::resource_id(&stream, fs::KIND_INPUT_STREAM).is_some() {
                return make_ready_pollable();
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::make_pollable(obj);
                }
            }
            make_ready_pollable()
        }),
    );

    // [resource-drop]input-stream(self) — release the registry entry for a
    // resource-backed stream. fd streams (stdin) carry no registry entry, so the
    // drop is a no-op; GC reclaims the handle object either way.
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[resource-drop]input-stream",
        Some((INPUT_STREAM, ResourceMemberKind::Destructor)),
        // `own`, not `borrow` — dropping CONSUMES the handle. This is the one
        // distinction the `[resource-drop]` name prefix could never carry.
        vec![ValType::Own(INPUT_STREAM.to_string())],
        vec![],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            if let Some(id) = fs::resource_id(&stream, fs::KIND_INPUT_STREAM) {
                fs::registry().lock().unwrap().input_streams.remove(&id);
            }
            Value::Null
        }),
    );
}

// ── wasi:io/streams — output-stream ──────────────────────────────────────────

fn register_output_stream(vm: &mut VM) {
    // [method]output-stream.check-write(self) -> result<u64, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.check-write",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(OUTPUT_STREAM)],
        vec![wasi_result(ValType::I64)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            if fs::resource_id(&stream, fs::KIND_OUTPUT_STREAM).is_some() {
                return Value::F64(65536.0);
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return Value::I64(65536);
                }
            }
            if matches!(stream_fd(&stream), Some(1) | Some(2)) {
                return Value::F64(65536.0);
            }
            fs::err("bad-descriptor")
        }),
    );

    // [method]output-stream.write(self, contents: list<u8>) -> result<_, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.write",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        // WIT's ok arm here is `_` — empty. `ValType` has no unit, so `Any`
        // stands in; what the declaration is really saying is "this can FAIL",
        // which is the half every caller has to handle.
        vec![borrowed(OUTPUT_STREAM), ValType::List(Box::new(ValType::I32))],
        vec![wasi_result(ValType::Any)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            let bytes = skt::bytes_from_value(args.get(1).unwrap_or(&Value::Null));
            if let Some(id) = fs::resource_id(&stream, fs::KIND_OUTPUT_STREAM) {
                return write_file_stream(id, &bytes, false);
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::write_stream_bytes(&obj, &bytes, false);
                }
            }
            if let Some(fd) = stream_fd(&stream) {
                let r = match fd {
                    1 => std::io::stdout().write_all(&bytes),
                    2 => std::io::stderr().write_all(&bytes),
                    _ => return fs::err("bad-descriptor"),
                };
                return match r {
                    Ok(_) => Value::Null,
                    Err(e) => fs::err(fs::map_io_error(&e)),
                };
            }
            fs::err("bad-descriptor")
        }),
    );

    // [method]output-stream.blocking-write-and-flush(self, contents: list<u8>) -> result<_, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.blocking-write-and-flush",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(OUTPUT_STREAM), ValType::List(Box::new(ValType::I32))],
        vec![wasi_result(ValType::Any)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            let bytes = skt::bytes_from_value(args.get(1).unwrap_or(&Value::Null));
            if let Some(id) = fs::resource_id(&stream, fs::KIND_OUTPUT_STREAM) {
                return write_file_stream(id, &bytes, true);
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::write_stream_bytes(&obj, &bytes, true);
                }
            }
            if let Some(fd) = stream_fd(&stream) {
                let r = match fd {
                    1 => std::io::stdout()
                        .write_all(&bytes)
                        .and_then(|_| std::io::stdout().flush()),
                    2 => std::io::stderr()
                        .write_all(&bytes)
                        .and_then(|_| std::io::stderr().flush()),
                    _ => return fs::err("bad-descriptor"),
                };
                return match r {
                    Ok(_) => Value::Null,
                    Err(e) => fs::err(fs::map_io_error(&e)),
                };
            }
            fs::err("bad-descriptor")
        }),
    );

    // [method]output-stream.flush(self) -> result<_, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.flush",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(OUTPUT_STREAM)],
        vec![wasi_result(ValType::Any)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            if let Some(id) = fs::resource_id(&stream, fs::KIND_OUTPUT_STREAM) {
                return flush_file_stream(id);
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::flush_stream(&obj);
                }
            }
            if let Some(fd) = stream_fd(&stream) {
                return match fd {
                    1 => {
                        let _ = std::io::stdout().flush();
                        Value::Null
                    }
                    2 => {
                        let _ = std::io::stderr().flush();
                        Value::Null
                    }
                    _ => fs::err("bad-descriptor"),
                };
            }
            fs::err("bad-descriptor")
        }),
    );

    // [method]output-stream.blocking-flush(self) -> result<_, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.blocking-flush",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(OUTPUT_STREAM)],
        vec![wasi_result(ValType::Any)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            if let Some(id) = fs::resource_id(&stream, fs::KIND_OUTPUT_STREAM) {
                return flush_file_stream(id);
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::flush_stream(&obj);
                }
            }
            if let Some(fd) = stream_fd(&stream) {
                return match fd {
                    1 => {
                        let _ = std::io::stdout().flush();
                        Value::Null
                    }
                    2 => {
                        let _ = std::io::stderr().flush();
                        Value::Null
                    }
                    _ => fs::err("bad-descriptor"),
                };
            }
            fs::err("bad-descriptor")
        }),
    );

    // [method]output-stream.subscribe(self) -> pollable
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.subscribe",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(OUTPUT_STREAM)],
        vec![ValType::Own(POLLABLE.to_string())],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            if fs::resource_id(&stream, fs::KIND_OUTPUT_STREAM).is_some() {
                return make_ready_pollable();
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::make_pollable(obj);
                }
            }
            // fd streams (stdout/stderr/stdin) are always ready
            make_ready_pollable()
        }),
    );

    // [method]output-stream.write-zeroes(self, len: u64) -> result<_, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.write-zeroes",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(OUTPUT_STREAM), ValType::I64],
        vec![wasi_result(ValType::Any)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            let len = u64_arg(args, 1) as usize;
            let zeros = vec![0u8; len.min(64 * 1024)];
            if let Some(id) = fs::resource_id(&stream, fs::KIND_OUTPUT_STREAM) {
                return write_file_stream(id, &zeros, false);
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::write_stream_bytes(&obj, &zeros, false);
                }
            }
            if let Some(fd) = stream_fd(&stream) {
                let r = match fd {
                    1 => std::io::stdout().write_all(&zeros),
                    2 => std::io::stderr().write_all(&zeros),
                    _ => return fs::err("bad-descriptor"),
                };
                return match r {
                    Ok(_) => Value::Null,
                    Err(e) => fs::err(fs::map_io_error(&e)),
                };
            }
            fs::err("bad-descriptor")
        }),
    );

    // [method]output-stream.blocking-write-zeroes-and-flush(self, len: u64) -> result<_, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.blocking-write-zeroes-and-flush",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        vec![borrowed(OUTPUT_STREAM), ValType::I64],
        vec![wasi_result(ValType::Any)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            let len = u64_arg(args, 1) as usize;
            let zeros = vec![0u8; len.min(64 * 1024)];
            if let Some(id) = fs::resource_id(&stream, fs::KIND_OUTPUT_STREAM) {
                return write_file_stream(id, &zeros, true);
            }
            if let Some(obj) = as_obj(&stream) {
                if skt::stream_socket_id(&obj).is_some() {
                    return skt::write_stream_bytes(&obj, &zeros, true);
                }
            }
            if let Some(fd) = stream_fd(&stream) {
                let r = match fd {
                    1 => std::io::stdout()
                        .write_all(&zeros)
                        .and_then(|_| std::io::stdout().flush()),
                    2 => std::io::stderr()
                        .write_all(&zeros)
                        .and_then(|_| std::io::stderr().flush()),
                    _ => return fs::err("bad-descriptor"),
                };
                return match r {
                    Ok(_) => Value::Null,
                    Err(e) => fs::err(fs::map_io_error(&e)),
                };
            }
            fs::err("bad-descriptor")
        }),
    );

    // [method]output-stream.splice(self, src: borrow<input-stream>, len: u64) -> result<u64, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.splice",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        // TWO stream handles, both borrowed and of DIFFERENT resource types.
        // Splicing consumes neither end — that is the fact a bare
        // `(Value, Value)` pair could not state.
        vec![
            borrowed(OUTPUT_STREAM),
            borrowed(INPUT_STREAM),
            ValType::I64,
        ],
        vec![wasi_result(ValType::I64)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let dst = args.first().cloned().unwrap_or(Value::Null);
            let src = args.get(1).cloned().unwrap_or(Value::Null);
            let len = u64_arg(args, 2) as usize;
            // Prefer socket splice when both ends are socket streams
            if let (Some(src_obj), Some(dst_obj)) = (as_obj(&src), as_obj(&dst)) {
                if skt::stream_socket_id(&src_obj).is_some()
                    && skt::stream_socket_id(&dst_obj).is_some()
                {
                    return skt::splice_streams(&src_obj, &dst_obj, len, false);
                }
            }
            let bytes = read_bytes_from_stream(&src, len, false);
            if bytes.is_empty() {
                return Value::I64(0);
            }
            let ok = write_bytes_to_stream(&dst, &bytes, false);
            if ok {
                Value::I64(bytes.len() as i64)
            } else {
                fs::err("io")
            }
        }),
    );

    // [method]output-stream.blocking-splice(self, src: borrow<input-stream>, len: u64) -> result<u64, stream-error>
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[method]output-stream.blocking-splice",
        Some((OUTPUT_STREAM, ResourceMemberKind::Method)),
        vec![
            borrowed(OUTPUT_STREAM),
            borrowed(INPUT_STREAM),
            ValType::I64,
        ],
        vec![wasi_result(ValType::I64)],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let dst = args.first().cloned().unwrap_or(Value::Null);
            let src = args.get(1).cloned().unwrap_or(Value::Null);
            let len = u64_arg(args, 2) as usize;
            if let (Some(src_obj), Some(dst_obj)) = (as_obj(&src), as_obj(&dst)) {
                if skt::stream_socket_id(&src_obj).is_some()
                    && skt::stream_socket_id(&dst_obj).is_some()
                {
                    return skt::splice_streams(&src_obj, &dst_obj, len, true);
                }
            }
            let bytes = read_bytes_from_stream(&src, len, true);
            if bytes.is_empty() {
                return Value::I64(0);
            }
            let ok = write_bytes_to_stream(&dst, &bytes, true);
            if ok {
                Value::I64(bytes.len() as i64)
            } else {
                fs::err("io")
            }
        }),
    );

    // [resource-drop]output-stream(self) — release the registry entry for a
    // resource-backed stream; fd streams (stdout/stderr) are a no-op.
    wasi_fn(
        vm,
        "wasi:io/streams",
        "[resource-drop]output-stream",
        Some((OUTPUT_STREAM, ResourceMemberKind::Destructor)),
        vec![ValType::Own(OUTPUT_STREAM.to_string())],
        vec![],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let stream = args.first().cloned().unwrap_or(Value::Null);
            if let Some(id) = fs::resource_id(&stream, fs::KIND_OUTPUT_STREAM) {
                fs::registry().lock().unwrap().output_streams.remove(&id);
            }
            Value::Null
        }),
    );
}

// ── wasi:io/poll ─────────────────────────────────────────────────────────────

fn register_pollable(vm: &mut VM) {
    // [method]pollable.ready(self) -> bool
    wasi_fn(
        vm,
        "wasi:io/poll",
        "[method]pollable.ready",
        Some((POLLABLE, ResourceMemberKind::Method)),
        vec![borrowed(POLLABLE)],
        vec![ValType::Bool],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(obj) = args.first().and_then(as_obj) else {
                return Value::Bool(false);
            };
            let direct_ready = obj.lock().unwrap().properties.get("__ready").cloned();
            if let Some(v) = direct_ready {
                return v;
            }
            Value::Bool(skt::pollable_ready(&obj))
        }),
    );

    // [method]pollable.block(self)
    wasi_fn(
        vm,
        "wasi:io/poll",
        "[method]pollable.block",
        Some((POLLABLE, ResourceMemberKind::Method)),
        // Confirmed against a real call site: emitted with argc 1, so the
        // handle IS passed even though this closure never reads it.
        vec![borrowed(POLLABLE)],
        vec![],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(obj) = args.first().and_then(as_obj) {
                skt::block_until_ready(&obj);
            }
            Value::Null
        }),
    );

    // poll(list<borrow<pollable>>) -> list<u32>  — unified handler also registered here
    wasi_fn(
        vm,
        "wasi:io/poll",
        "poll",
        // No resource: `poll` is a FREE function over a list of pollables, not
        // a member of one. It is the only function in this file that is not.
        None,
        vec![ValType::List(Box::new(borrowed(POLLABLE)))],
        vec![ValType::List(Box::new(ValType::I32))],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let Some(list) = args.first() else {
                return skt::value_array(vec![]);
            };
            let Some(pollables) = skt::value_array_elements(list) else {
                return skt::value_array(vec![]);
            };
            let mut ready = collect_ready_indices(&pollables);
            if ready.is_empty() {
                let start = Instant::now();
                while ready.is_empty() && start.elapsed() < Duration::from_secs(1) {
                    thread::sleep(Duration::from_millis(1));
                    ready = collect_ready_indices(&pollables);
                }
            }
            skt::value_array(ready.into_iter().map(|i| Value::I32(i as i32)).collect())
        }),
    );

    // [resource-drop]pollable(self) — pollables are transient handle objects with
    // no registry backing, so the drop is a GC-safe no-op.
    wasi_fn(
        vm,
        "wasi:io/poll",
        "[resource-drop]pollable",
        Some((POLLABLE, ResourceMemberKind::Destructor)),
        vec![ValType::Own(POLLABLE.to_string())],
        vec![],
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Null),
    );
}

// ── wasi:io/error ─────────────────────────────────────────────────────────────

fn register_error(vm: &mut VM) {
    // [method]error.to-debug-string(self) -> string
    wasi_fn(
        vm,
        "wasi:io/error",
        "[method]error.to-debug-string",
        Some((ERROR, ResourceMemberKind::Method)),
        vec![borrowed(ERROR)],
        vec![ValType::String],
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let inner = obj.lock().unwrap();
                if let Some(code) = inner.properties.get("__wasi_error") {
                    return Value::String(Arc::from(format!("wasi-io-error: {}", code).as_str()));
                }
                let parts: Vec<String> = inner
                    .properties
                    .iter()
                    .filter(|(k, _)| !k.starts_with("__"))
                    .map(|(k, v)| format!("{}={}", k, v))
                    .collect();
                return Value::String(Arc::from(
                    format!("io-error({})", parts.join(", ")).as_str(),
                ));
            }
            Value::String(Arc::from("io-error(unknown)"))
        }),
    );

    // [resource-drop]error(self) — error records are plain handle objects with no
    // registry backing, so the drop is a GC-safe no-op.
    wasi_fn(
        vm,
        "wasi:io/error",
        "[resource-drop]error",
        Some((ERROR, ResourceMemberKind::Destructor)),
        vec![ValType::Own(ERROR.to_string())],
        vec![],
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Null),
    );
}
