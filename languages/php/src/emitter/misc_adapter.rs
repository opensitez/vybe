use std::sync::Arc;
use vybe_compiler::primitives::instructions::core_wasm;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::object::{
    emit_bind_bound_method, emit_bind_getter, emit_bind_setter,
};

const SERIAL_KIND_KEY: &str = "vybe$php_ser_kind";

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_op(Op::NULL, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

fn call_import_into(
    _imports: &mut Chunk,
    code: &mut Chunk,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = code.add_import(module.to_string(), name.to_string());
    code.emit_call(idx, argc, line);
}

fn call_ref(chunk: &mut Chunk, argc: u8, line: u32) {
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(argc, line);
}

fn ref_func(chunk: &mut Chunk, func_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::REF_FUNC, func_idx as u16, line);
    chunk.emit(0, line);
}

fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_GET, idx, line);
}

fn struct_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn dynamic_get_from_slots(chunk: &mut Chunk, obj_slot: u16, key_slot: u16, line: u32) {
    lget(chunk, obj_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn dynamic_set_from_slots(
    chunk: &mut Chunk,
    obj_slot: u16,
    key_slot: u16,
    value_slot: u16,
    line: u32,
) {
    lget(chunk, obj_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
}

fn set_struct_from_slot(chunk: &mut Chunk, obj_slot: u16, key: &str, value_slot: u16, line: u32) {
    lget(chunk, obj_slot, line);
    lget(chunk, value_slot, line);
    struct_set_key(chunk, key, line);
}

/// Apply one PHP `header()` line through Node's response API.
///
/// PHP hands `header()` a RAW field line — `"Content-Type: text/plain"`, or a
/// status line `"HTTP/1.1 404 Not Found"`. Splitting that is PHP's semantics,
/// so it happens here. Node's response object has `setHeader`, `appendHeader`
/// and `statusCode` and nothing that takes a raw line, so a `send_header_raw`
/// host function was PHP's `header()` wearing a Node badge.
fn emit_apply_header(
    chunks: &mut [Chunk],
    current: usize,
    header_slot: u16,
    replace_slot: Option<u16>,
    line: u32,
) {
    let set_status = chunks[0].add_import("node:http".to_string(), "set_status".to_string());
    let set_header = chunks[0].add_import("node:http".to_string(), "set_header".to_string());
    let add_header = chunks[0].add_import("node:http".to_string(), "add_header".to_string());
    let index_of = chunks[0].add_import("ecma:string".to_string(), "indexOf".to_string());
    let starts_with = chunks[0].add_import("ecma:string".to_string(), "startsWith".to_string());
    let substring = chunks[0].add_import("wasm:js-string".to_string(), "substring".to_string());
    let length = chunks[0].add_import("wasm:js-string".to_string(), "length".to_string());
    let trim = chunks[0].add_import("ecma:string".to_string(), "trim".to_string());
    let to_number = chunks[0].add_import("ecma:value".to_string(), "toNumber".to_string());

    let at = alloc_local(&mut chunks[current]);
    let rest = alloc_local(&mut chunks[current]);
    let name_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];

    lget(chunk, header_slot, line);
    push_str(chunk, "HTTP/", line);
    chunk.emit_call(starts_with, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    // Status line: the code is the token between the first two spaces.
    lget(chunk, header_slot, line);
    push_str(chunk, " ", line);
    chunk.emit_call(index_of, 2, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, at, line);

    lget(chunk, header_slot, line);
    lget(chunk, at, line);
    lget(chunk, header_slot, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_call(substring, 3, line);
    lset(chunk, rest, line);

    lget(chunk, rest, line);
    push_str(chunk, " ", line);
    chunk.emit_call(index_of, 2, line);
    lset(chunk, at, line);

    lget(chunk, at, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    lget(chunk, rest, line);
    chunk.emit_i32_const(0, line);
    lget(chunk, at, line);
    chunk.emit_call(substring, 3, line);
    chunk.emit_else(line);
    lget(chunk, rest, line);
    chunk.emit_end(line);
    chunk.emit_call(to_number, 1, line);
    chunk.emit_call(set_status, 1, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_else(line);

    // Field line: `Name: value`, split at the FIRST colon so values keeping
    // their own colons (`Location: http://…`) survive intact.
    lget(chunk, header_slot, line);
    push_str(chunk, ":", line);
    chunk.emit_call(index_of, 2, line);
    lset(chunk, at, line);

    lget(chunk, at, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);

    lget(chunk, header_slot, line);
    chunk.emit_i32_const(0, line);
    lget(chunk, at, line);
    chunk.emit_call(substring, 3, line);
    chunk.emit_call(trim, 1, line);
    lset(chunk, name_slot, line);

    lget(chunk, name_slot, line);
    lget(chunk, header_slot, line);
    lget(chunk, at, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    lget(chunk, header_slot, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_call(substring, 3, line);
    chunk.emit_call(trim, 1, line);

    // `replace` defaults to true — PHP replaces a same-named header unless
    // told otherwise, which is `setHeader` vs `appendHeader` in Node.
    match replace_slot {
        Some(slot) => {
            lget(chunk, slot, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_if(line);
            chunk.emit_call(set_header, 2, line);
            chunk.emit_else(line);
            chunk.emit_call(add_header, 2, line);
            chunk.emit_end(line);
        }
        None => chunk.emit_call(set_header, 2, line),
    }
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// PHP `php_sapi_name()` — which SAPI this script is running under.
///
/// A SAPI is a PHP concept; Node has no notion of one, so a `php_sapi_name`
/// host function was PHP's function name sitting in `node:http`. Whether a
/// request is in flight is answered by the shared request primitive, the same
/// source `$_SERVER` reads.
pub fn emit_php_sapi_name(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let request = chunk.add_constant(Value::String(Arc::from(
        vybe_compiler::primitives::http_request_env::REQUEST_GLOBAL,
    )));
    chunk.emit_op_u16(Op::GLOBAL_GET, request, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_str(chunk, "cli", line);
    chunk.emit_else(line);
    push_str(chunk, "vybex-server", line);
    chunk.emit_end(line);
}

/// PHP `http_response_code()` — get with no argument, set with one.
///
/// Arity-based get/set is PHP's calling convention, not Node's. Node has a
/// `statusCode` property, which is `node:http`'s `status` / `set_status`; a
/// combined `http_response_code` host function was PHP's function name sitting
/// in Node's namespace.
pub fn emit_php_http_response_code(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let status = chunks[0].add_import("node:http".to_string(), "status".to_string());
    let set_status = chunks[0].add_import("node:http".to_string(), "set_status".to_string());
    let chunk = &mut chunks[current];

    // PHP answers 200 when nothing has set a status; Node's `statusCode`
    // reports 0 for "unset". The default is PHP's, so it is applied here — the
    // old host function baked it into `node:http` instead.
    if argc == 0 {
        emit_status_or_default(chunk, status, line);
        return;
    }

    // Setting: `http_response_code(404)`. A zero or absent code is a read, the
    // same way the host function behaved.
    let code = alloc_local(chunk);
    lset(chunk, code, line);
    lget(chunk, code, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, code, line);
    chunk.emit_call(set_status, 1, line);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, code, line);
    chunk.emit_else(line);
    emit_status_or_default(chunk, status, line);
    chunk.emit_end(line);
}

/// Node's `statusCode`, or PHP's 200 default when nothing has set one.
fn emit_status_or_default(chunk: &mut Chunk, status: u16, line: u32) {
    let current = alloc_local(chunk);
    chunk.emit_call(status, 0, line);
    lset(chunk, current, line);
    lget(chunk, current, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, current, line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(200.0), line);
    chunk.emit_end(line);
}

pub fn emit_php_header(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let set_status_import = chunks[0].add_import("node:http".to_string(), "set_status".to_string());
    let status_import = chunks[0].add_import("node:http".to_string(), "status".to_string());
    let string_import = chunks[0].add_import("ecma:string".to_string(), "String".to_string());
    let lower_import = chunks[0].add_import("ecma:string".to_string(), "toLowerCase".to_string());
    let starts_with_import =
        chunks[0].add_import("ecma:string".to_string(), "startsWith".to_string());
    let chunk = &mut chunks[current];

    let header_slot = alloc_local(chunk);
    let replace_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let response_code_slot = if argc >= 3 {
        Some(alloc_local(chunk))
    } else {
        None
    };

    if let Some(slot) = response_code_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = replace_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, header_slot, line);

    emit_apply_header(chunks, current, header_slot, replace_slot, line);

    // `header($h, $replace, $code)` — an explicit third argument sets the
    // status outright.
    let chunk = &mut chunks[current];
    if let Some(slot) = response_code_slot {
        lget(chunk, slot, line);
        push_const(chunk, Value::F64(0.0), line);
        vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
        chunk.emit_if(line);
        lget(chunk, slot, line);
        chunk.emit_call(set_status_import, 1, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);
    }

    if let Some(slot) = response_code_slot {
        lget(chunk, slot, line);
        push_const(chunk, Value::F64(0.0), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
    }

    lget(chunk, header_slot, line);
    chunk.emit_call(string_import, 1, line);
    chunk.emit_call(lower_import, 1, line);
    push_str(chunk, "location:", line);
    chunk.emit_call(starts_with_import, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_call(status_import, 0, line);
    push_const(chunk, Value::F64(200.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);

    push_const(chunk, Value::F64(302.0), line);
    chunk.emit_call(set_status_import, 1, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    if response_code_slot.is_some() {
        chunk.emit_end(line);
    }

    chunk.emit_op(Op::NULL, line);
}

pub fn emit_php_extension_loaded(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let string_import = chunks[0].add_import("ecma:string".to_string(), "String".to_string());
    let lower_import = chunks[0].add_import("ecma:string".to_string(), "toLowerCase".to_string());
    let chunk = &mut chunks[current];
    if argc == 0 {
        push_const(chunk, Value::Bool(false), line);
        return;
    }

    // extension_loaded() is unary; discard extra args defensively.
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }

    chunk.emit_call(string_import, 1, line);
    chunk.emit_call(lower_import, 1, line);

    let ext_slot = alloc_local(chunk);
    lset(chunk, ext_slot, line);

    lget(chunk, ext_slot, line);
    push_str(chunk, "mysqli", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, ext_slot, line);
    push_str(chunk, "mysqlnd", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, ext_slot, line);
    push_str(chunk, "pdo_mysql", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, ext_slot, line);
    push_str(chunk, "mysql", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_php_phpversion(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "8.0.0", line);
}

/// PHP `phpinfo([$flags])` — text report of core runtime facts (subset of full
/// PHP `phpinfo()`). Writes to stdout without an extra trailing newline beyond
/// the final report line, returns `true`.
pub fn emit_php_phpinfo(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let write_idx = chunk.add_import("wasi:cli/stdout", "write-via-stream");
    let rd_slot = alloc_local(chunk);
    let wr_slot = alloc_local(chunk);
    vybe_compiler::primitives::io::emit_write_stdout_with_imports(
        chunk,
        write_idx,
        rd_slot,
        wr_slot,
        line,
        |c| {
            push_str(
                c,
                "phpinfo()\n\
                 PHP Version => 8.0.0\n\
                 System => Darwin\n\
                 Build Date => vybe\n\
                 Server API => cli\n\
                 PHP API => vybex\n\
                 PHP Extension Build => vybe\n\
                 Zend Extension Build => n/a\n\
                 PHP Integer Size => 8\n",
                line,
            );
        },
    );
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_session_start(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let map_new_import = chunks[0].add_import("ecma:map".to_string(), "new".to_string());
    let chunk = &mut chunks[current];
    let needs_cookie = chunk.add_constant(Value::String(Arc::from("__php_session_needs_cookie")));
    let session_id = chunk.add_constant(Value::String(Arc::from("__php_session_id")));
    let started = chunk.add_constant(Value::String(Arc::from("__php_session_started")));
    let destroyed = chunk.add_constant(Value::String(Arc::from("__php_session_destroyed")));
    let session = chunk.add_constant(Value::String(Arc::from("$_SESSION")));

    push_const(chunk, Value::Bool(true), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, started, line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, destroyed, line);

    // Mirror the lifecycle into the SHARED session primitive's global so
    // `session_status()` — which dispatches to `common:http_session.status` —
    // reports ACTIVE. Without this it kept answering "not started" right after
    // a successful `session_start()`.
    let status = chunk.add_constant(Value::String(Arc::from(
        vybe_compiler::primitives::http_session::SESSION_STATUS_GLOBAL,
    )));
    push_const(
        chunk,
        Value::I32(vybe_compiler::primitives::http_session::STATUS_ACTIVE),
        line,
    );
    chunk.emit_op_u16(Op::GLOBAL_SET, status, line);

    chunk.emit_op_u16(Op::GLOBAL_GET, session, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_call(map_new_import, 0, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, session, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::GLOBAL_GET, needs_cookie, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    // The session id goes out RAW. Real PHP does not URL-encode session ids,
    // and an encoded cookie would disagree with what `session_id()` returns —
    // the value apps compare against and put in URLs.
    push_str(chunk, "PHPSESSID", line);
    chunk.emit_op_u16(Op::GLOBAL_GET, session_id, line);
    // No expiry: a session cookie lives until the browser closes.
    vybe_compiler::primitives::http_cookie::emit_serialize(chunks, current, 2, line);
    emit_send_cookie(chunks, current, line);

    let chunk = &mut chunks[current];
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, needs_cookie, line);
    chunk.emit_end(line);
    push_const(chunk, Value::Bool(true), line);
}

/// Send the serialized cookie on the stack as a `Set-Cookie` header.
/// Stack: `[cookie]` → `[]`.
///
/// `add_header` is Node's `appendHeader`. Every other header replaces a
/// same-named one; `Set-Cookie` is the header that cannot be combined, so each
/// cookie is its own field line and replacing would silently drop all but the
/// last (RFC 6265 §3.1, RFC 9110 §5.3).
///
/// The cookie itself is built by `primitives/http_cookie.rs`. Nothing here
/// spells a cookie by hand: `session_start` and `session_destroy` used to
/// concatenate their own `Set-Cookie: PHPSESSID=…` literals while
/// `setcookie()` next door went through the shared serializer.
fn emit_send_cookie(chunks: &mut [Chunk], current: usize, line: u32) {
    let add_header = chunks[0].add_import("node:http".to_string(), "add_header".to_string());
    let cookie_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, cookie_slot, line);
    push_str(chunk, "Set-Cookie", line);
    lget(chunk, cookie_slot, line);
    chunk.emit_call(add_header, 2, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_php_session_unset(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    let session = chunk.add_constant(Value::String(Arc::from("$_SESSION")));
    chunk.emit_op_u16(Op::GLOBAL_SET, session, line);
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_session_destroy(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_php_session_unset(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    // Close the session in the shared primitive too, so `session_status()`
    // stops reporting ACTIVE.
    let status = chunk.add_constant(Value::String(Arc::from(
        vybe_compiler::primitives::http_session::SESSION_STATUS_GLOBAL,
    )));
    push_const(
        chunk,
        Value::I32(vybe_compiler::primitives::http_session::STATUS_NONE),
        line,
    );
    chunk.emit_op_u16(Op::GLOBAL_SET, status, line);
    let started = chunk.add_constant(Value::String(Arc::from("__php_session_started")));
    let destroyed = chunk.add_constant(Value::String(Arc::from("__php_session_destroyed")));
    let needs_cookie = chunk.add_constant(Value::String(Arc::from("__php_session_needs_cookie")));

    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, started, line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, destroyed, line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_op_u16(Op::GLOBAL_SET, needs_cookie, line);

    // Deleting a cookie means expiring it in the past (RFC 6265 §3.1) — an
    // empty value alone leaves the cookie in the jar. The date formatting is
    // the shared serializer's (`expires_becomes_an_http_date`), not a literal
    // spelled out here.
    push_str(chunk, "PHPSESSID", line);
    push_str(chunk, "", line);
    let attrs = chunks[0].add_import("ecma:map".to_string(), "new".to_string());
    chunks[current].emit_call(attrs, 0, line);
    let attrs_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, attrs_slot, line);
    lget(chunk, attrs_slot, line);
    push_str(chunk, "expires", line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], attrs_slot, line);
    vybe_compiler::primitives::http_cookie::emit_serialize(chunks, current, 3, line);
    emit_send_cookie(chunks, current, line);

    push_const(&mut chunks[current], Value::Bool(true), line);
}

fn helper_loop_start(chunk: &mut Chunk, line: u32) -> vybe_compiler::primitives::loops::LoopState {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    vybe_compiler::primitives::loops::LoopState {
        block_patch,
        loop_patch,
        body_block_patch: None,
    }
}

fn helper_loop_cond(chunk: &mut Chunk, line: u32) {
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);
}

fn helper_loop_end(chunk: &mut Chunk, state: vybe_compiler::primitives::loops::LoopState, line: u32) {
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(state.loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(state.block_patch);
}

fn emit_nullish_return(chunk: &mut Chunk, value_slot: u16, line: u32) {
    lget(chunk, value_slot, line);
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_else(line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    chunk.emit_if(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op(Op::RETURN, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_is_array_into(imports: &mut Chunk, code: &mut Chunk, value_slot: u16, line: u32) {
    lget(code, value_slot, line);
    call_import_into(imports, code, "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(code, line);
}

fn bump_loop_index(chunk: &mut Chunk, i_slot: u16, line: u32) {
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
}

fn build_php_alloc_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let types = chunks[0].types.clone();

    let mut helper = create_function_chunk("__php_unserialize_alloc", 1);
    helper.alloc_scratch(1);
    let class_slot = 0;
    let obj_slot = alloc_local(&mut helper);

    {
        let imports = &mut chunks[0];
        emit_nullish_return(&mut helper, class_slot, line);

        for ty in types.iter().filter(|ty| !ty.is_interface) {
            lget(&mut helper, class_slot, line);
            push_str(&mut helper, &ty.name, line);
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
            helper.emit_if(line);

            helper.emit_op_u16(Op::STRUCT_NEW, 0, line);
            lset(&mut helper, obj_slot, line);

            lget(&mut helper, obj_slot, line);
            push_str(&mut helper, &ty.name, line);
            struct_set_key(&mut helper, "__type", line);

            lget(&mut helper, obj_slot, line);
            push_str(&mut helper, &ty.name.to_lowercase(), line);
            struct_set_key(&mut helper, "__control_name", line);

            let tid_name =
                helper.add_constant(Value::String(Arc::from(format!("__tid_{}", ty.name))));
            lget(&mut helper, obj_slot, line);
            helper.emit_op_u16(Op::GLOBAL_GET, tid_name, line);
            let tid_key = helper.add_constant(Value::String(Arc::from("__type_id")));
            helper.emit_op_u16(Op::STRUCT_SET, tid_key, line);
            helper.emit_op(Op::DROP, line);

            for field in &ty.fields {
                lget(&mut helper, obj_slot, line);
                helper.emit_op(Op::NULL, line);
                struct_set_key(&mut helper, field, line);
            }

            for (method_name, method_chunk_idx) in &ty.methods {
                if method_name.starts_with("__get_") {
                    let prop = method_name
                        .strip_prefix("__get_")
                        .unwrap_or(method_name.as_str());
                    emit_bind_getter(&mut helper, obj_slot, prop, *method_chunk_idx, line);
                } else if method_name.starts_with("__set_") {
                    let prop = method_name
                        .strip_prefix("__set_")
                        .unwrap_or(method_name.as_str());
                    emit_bind_setter(&mut helper, obj_slot, prop, *method_chunk_idx, line);
                } else {
                    emit_bind_bound_method(
                        &mut helper,
                        obj_slot,
                        method_name,
                        *method_chunk_idx,
                        None,
                        false, // PHP binds the receiver at call time, not on access
                        line,
                    );
                }
            }

            lget(&mut helper, obj_slot, line);
            helper.emit_op(Op::RETURN, line);
            helper.emit_end(line);
        }

        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        helper.emit_op(Op::RETURN, line);
    }

    chunks.push(helper);
    helper_idx
}

/// Recursive PHP value → JSON-serializable shape. An associative array is an
/// `ObjectKind::Map`, and `ecma:json.stringify` renders a bare Map as `{}` (ECMA
/// §25 — `JSON.stringify(new Map())` is `{}`), so PHP must convert Map → plain
/// Object first. Arrays recurse on elements; nested Maps recurse too. Key order
/// is the Map's native (`ecma:object.keys`) insertion order — no `__keys`/CSV
/// side-band. The helper self-recurses via its own func ref.
#[allow(dead_code)]
fn build_php_json_normalize_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut helper = create_function_chunk("__php_json_normalize", 1);
    helper.alloc_scratch(1);

    let value_slot = 0u16;
    let out_slot = alloc_local(&mut helper);
    let keys_slot = alloc_local(&mut helper);
    let key_slot = alloc_local(&mut helper);
    let i_slot = alloc_local(&mut helper);
    let n_slot = alloc_local(&mut helper);
    let _type_slot = alloc_local(&mut helper);

    {
        let imports = &mut chunks[0];
        // ALL imports in a helper chunk must register on `imports` (chunks[0])
        // via the `_into` ops — `add_import` is per-chunk and emit_call
        // resolves against chunks[0]'s table, so the non-`_into` ops (which use
        // the helper's own list) produce clashing indices.

        // null / undefined → pass through (stringify handles them).
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        lget(&mut helper, value_slot, line);
        {
            let undef_idx = helper.add_import("wasm:js-undefined", "test");
            helper.emit_call(undef_idx, 1, line);
        }
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        // Sequential array → new array of normalized elements.
        lget(&mut helper, value_slot, line);
        call_import_into(imports, &mut helper, "ecma:array", "isArray", 1, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool_into(imports, &mut helper, line);
        helper.emit_if(line);
        helper.emit_array_new_fixed(0, 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let arr_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt_into(imports, &mut helper, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool_into(imports, &mut helper, line);
        vybe_compiler::primitives::ops::emit_dyn_not_into(imports, &mut helper, line);
        helper.emit_br_if(1, line);
        lget(&mut helper, out_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, value_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        call_ref(&mut helper, 1, line);
        call_import_into(imports, &mut helper, "ecma:array", "push", 2, line);
        helper.emit_op(Op::DROP, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, arr_loop, line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        // Object / Map → fromEntries of [k, normalize(v[k])] in native key order.
        // object test: not null AND not number AND not string AND not boolean
        let obj_norm_slot = alloc_local(&mut helper);
        lget(&mut helper, value_slot, line);
        helper.emit_op_u16(Op::LOCAL_SET, obj_norm_slot, line);
        // not null
        lget(&mut helper, obj_norm_slot, line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_op(Op::I32_EQZ, line);
        // AND not number
        lget(&mut helper, obj_norm_slot, line);
        let test_num_norm = helper.add_import("wasm:js-number", "test");
        helper.emit_call(test_num_norm, 1, line);
        helper.emit_op(Op::I32_EQZ, line);
        helper.emit_op(Op::I32_AND, line);
        // AND not string
        lget(&mut helper, obj_norm_slot, line);
        let test_str_norm = helper.add_import("wasm:js-string", "test");
        helper.emit_call(test_str_norm, 1, line);
        helper.emit_op(Op::I32_EQZ, line);
        helper.emit_op(Op::I32_AND, line);
        // AND not boolean
        lget(&mut helper, obj_norm_slot, line);
        let test_bool_norm = helper.add_import("wasm:js-boolean", "test");
        helper.emit_call(test_bool_norm, 1, line);
        helper.emit_op(Op::I32_EQZ, line);
        helper.emit_op(Op::I32_AND, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "keys", 1, line);
        lset(&mut helper, keys_slot, line);
        helper.emit_array_new_fixed(0, 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, keys_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let obj_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt_into(imports, &mut helper, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool_into(imports, &mut helper, line);
        vybe_compiler::primitives::ops::emit_dyn_not_into(imports, &mut helper, line);
        helper.emit_br_if(1, line);
        lget(&mut helper, keys_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        // pair = [ key, normalize(value[key]) ] ; out.push(pair)
        lget(&mut helper, out_slot, line);
        lget(&mut helper, key_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, value_slot, line);
        lget(&mut helper, key_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        call_ref(&mut helper, 1, line);
        helper.emit_array_new_fixed(0, 2, line);
        call_import_into(imports, &mut helper, "ecma:array", "push", 2, line);
        helper.emit_op(Op::DROP, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, obj_loop, line);
        lget(&mut helper, out_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "fromEntries", 1, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        // Primitive (boolean / number / string) → pass through.
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
    }

    chunks.push(helper);
    helper_idx
}

/// Build the normalizer and call it on `value_slot`, leaving the
/// JSON-serializable (Map-free) value on the stack.
pub fn emit_php_json_normalize(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    line: u32,
) {
    // Delegate to the shared `vybe_compiler::primitives::json` normalizer. PHP serializes
    // every object's own enumerable keys (props mode), applies no encoder hook,
    // and preserves native key order — byte-compatible with the former
    // hand-rolled helper, now unified with Python's json path.
    let (default_slot, sort_slot, props_slot) = {
        let c = &mut chunks[current];
        (c.alloc_scratch(1), c.alloc_scratch(1), c.alloc_scratch(1))
    };
    {
        let c = &mut chunks[current];
        c.emit_op(Op::NULL, line); // default (no encoder hook)
        lset(c, default_slot, line);
        c.emit_bool_const(false, line); // sort_keys
        lset(c, sort_slot, line);
        c.emit_bool_const(true, line); // serialize props
        lset(c, props_slot, line);
    }
    vybe_compiler::primitives::json::emit_normalize(
        chunks,
        current,
        value_slot,
        default_slot,
        sort_slot,
        props_slot,
        line,
    );
}

fn build_php_serialize_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut helper = create_function_chunk("__php_serialize_value", 1);
    helper.alloc_scratch(1);

    let value_slot = 0;
    let _type_slot = alloc_local(&mut helper);
    let out_slot = alloc_local(&mut helper);
    let items_slot = alloc_local(&mut helper);
    let assoc_slot = alloc_local(&mut helper);
    let names_slot = alloc_local(&mut helper);
    let key_slot = alloc_local(&mut helper);
    let tmp_slot = alloc_local(&mut helper);
    let i_slot = alloc_local(&mut helper);
    let n_slot = alloc_local(&mut helper);
    let method_slot = alloc_local(&mut helper);

    {
        let imports = &mut chunks[0];
        emit_nullish_return(&mut helper, value_slot, line);

        // boolean test
        lget(&mut helper, value_slot, line);
        let test_bool_ser = helper.add_import("wasm:js-boolean", "test");
        helper.emit_call(test_bool_ser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        // number test
        lget(&mut helper, value_slot, line);
        let test_num_ser = helper.add_import("wasm:js-number", "test");
        helper.emit_call(test_num_ser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        // string test
        lget(&mut helper, value_slot, line);
        let test_str_ser = helper.add_import("wasm:js-string", "test");
        helper.emit_call(test_str_ser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        emit_is_array_into(imports, &mut helper, value_slot, line);
        helper.emit_if(line);

        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, out_slot, line);
        push_str(&mut helper, "array", line);
        struct_set_key(&mut helper, SERIAL_KIND_KEY, line);

        helper.emit_array_new_fixed(0, 0, line);
        lset(&mut helper, items_slot, line);
        lget(&mut helper, value_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let items_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, items_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, value_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        call_ref(&mut helper, 1, line);
        call_import_into(imports, &mut helper, "ecma:array", "push", 2, line);
        helper.emit_op(Op::DROP, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, items_loop, line);
        set_struct_from_slot(&mut helper, out_slot, "items", items_slot, line);

        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "vybe$assoc_keys_csv", line);
        lset(&mut helper, tmp_slot, line);
        lget(&mut helper, tmp_slot, line);
        helper.emit_dup(line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        helper.emit_op(Op::DROP, line);
        helper.emit_else(line);
        {
            let undef_idx = helper.add_import("wasm:js-undefined", "test");
            helper.emit_call(undef_idx, 1, line);
        }
        helper.emit_if(line);
        helper.emit_else(line);
        lget(&mut helper, tmp_slot, line);
        push_str(&mut helper, "\x1F", line);
        call_import_into(imports, &mut helper, "ecma:string", "split", 2, line);
        lset(&mut helper, names_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, assoc_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let assoc_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        ref_func(&mut helper, helper_idx, line);
        dynamic_get_from_slots(&mut helper, value_slot, key_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, assoc_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, assoc_loop, line);
        set_struct_from_slot(&mut helper, out_slot, "assoc", assoc_slot, line);
        helper.emit_end(line);
        helper.emit_end(line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);

        helper.emit_end(line);

        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__serialize", line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_ser = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_ser, line);
            lget(&mut helper, fn_slot_ser, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_ser, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_ser, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_ser, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, out_slot, line);
        push_str(&mut helper, "custom_object", line);
        struct_set_key(&mut helper, SERIAL_KIND_KEY, line);
        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__type", line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "class", tmp_slot, line);
        lget(&mut helper, method_slot, line);
        lget(&mut helper, value_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, tmp_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "payload", tmp_slot, line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__sleep", line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_slp = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_slp, line);
            lget(&mut helper, fn_slot_slp, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_slp, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_slp, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_slp, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, out_slot, line);
        push_str(&mut helper, "sleep_object", line);
        struct_set_key(&mut helper, SERIAL_KIND_KEY, line);
        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__type", line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "class", tmp_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, assoc_slot, line);
        lget(&mut helper, method_slot, line);
        lget(&mut helper, value_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, names_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let sleep_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        dynamic_get_from_slots(&mut helper, value_slot, key_slot, line);
        lset(&mut helper, tmp_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, tmp_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, assoc_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, sleep_loop, line);
        set_struct_from_slot(&mut helper, out_slot, "fields", assoc_slot, line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, out_slot, line);
        push_str(&mut helper, "object", line);
        struct_set_key(&mut helper, SERIAL_KIND_KEY, line);
        lget(&mut helper, value_slot, line);
        struct_get_key(&mut helper, "__type", line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "class", tmp_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "new", 0, line);
        lset(&mut helper, assoc_slot, line);
        lget(&mut helper, value_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "keys", 1, line);
        lset(&mut helper, names_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let object_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);

        for internal_key in [
            "__type",
            "__types",
            "__control_name",
            "__super",
            "vybe$assoc_keys_csv",
        ] {
            lget(&mut helper, key_slot, line);
            push_str(&mut helper, internal_key, line);
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
            helper.emit_if(line);
            bump_loop_index(&mut helper, i_slot, line);
            helper.emit_br(1, line);
            helper.emit_end(line);
        }

        dynamic_get_from_slots(&mut helper, value_slot, key_slot, line);
        lset(&mut helper, tmp_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_tmp = alloc_local(&mut helper);
            lget(&mut helper, tmp_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_tmp, line);
            lget(&mut helper, fn_slot_tmp, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_tmp, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_tmp, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_tmp, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        bump_loop_index(&mut helper, i_slot, line);
        helper.emit_br(1, line);
        helper.emit_end(line);

        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, tmp_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, assoc_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, object_loop, line);
        set_struct_from_slot(&mut helper, out_slot, "fields", assoc_slot, line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
    }

    chunks.push(helper);
    helper_idx
}

fn build_php_unserialize_helper(chunks: &mut Vec<Chunk>, alloc_idx: usize, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut helper = create_function_chunk("__php_unserialize_value", 1);
    helper.alloc_scratch(1);

    let node_slot = 0;
    let _type_slot = alloc_local(&mut helper);
    let kind_slot = alloc_local(&mut helper);
    let out_slot = alloc_local(&mut helper);
    let items_slot = alloc_local(&mut helper);
    let assoc_slot = alloc_local(&mut helper);
    let fields_slot = alloc_local(&mut helper);
    let names_slot = alloc_local(&mut helper);
    let key_slot = alloc_local(&mut helper);
    let tmp_slot = alloc_local(&mut helper);
    let i_slot = alloc_local(&mut helper);
    let n_slot = alloc_local(&mut helper);
    let method_slot = alloc_local(&mut helper);

    {
        let imports = &mut chunks[0];
        emit_nullish_return(&mut helper, node_slot, line);

        // boolean test
        lget(&mut helper, node_slot, line);
        let test_bool_unser = helper.add_import("wasm:js-boolean", "test");
        helper.emit_call(test_bool_unser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        // number test
        lget(&mut helper, node_slot, line);
        let test_num_unser = helper.add_import("wasm:js-number", "test");
        helper.emit_call(test_num_unser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);
        // string test
        lget(&mut helper, node_slot, line);
        let test_str_unser = helper.add_import("wasm:js-string", "test");
        helper.emit_call(test_str_unser, 1, line);
        helper.emit_if(line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, SERIAL_KIND_KEY, line);
        lset(&mut helper, kind_slot, line);
        lget(&mut helper, kind_slot, line);
        helper.emit_dup(line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        helper.emit_op(Op::DROP, line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_else(line);
        {
            let undef_idx = helper.add_import("wasm:js-undefined", "test");
            helper.emit_call(undef_idx, 1, line);
        }
        helper.emit_if(line);
        lget(&mut helper, node_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_else(line);

        lget(&mut helper, kind_slot, line);
        push_str(&mut helper, "array", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
        helper.emit_if(line);
        helper.emit_array_new_fixed(0, 0, line);
        lset(&mut helper, out_slot, line);
        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "items", line);
        lset(&mut helper, items_slot, line);
        lget(&mut helper, items_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let items_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, items_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        lget(&mut helper, out_slot, line);
        lget(&mut helper, tmp_slot, line);
        call_import_into(imports, &mut helper, "ecma:array", "push", 2, line);
        helper.emit_op(Op::DROP, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, items_loop, line);

        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "assoc", line);
        lset(&mut helper, assoc_slot, line);
        lget(&mut helper, assoc_slot, line);
        helper.emit_dup(line);
        helper.emit_op(Op::REF_IS_NULL, line);
        helper.emit_if(line);
        helper.emit_op(Op::DROP, line);
        helper.emit_else(line);
        {
            let undef_idx = helper.add_import("wasm:js-undefined", "test");
            helper.emit_call(undef_idx, 1, line);
        }
        helper.emit_if(line);
        helper.emit_else(line);
        lget(&mut helper, assoc_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "keys", 1, line);
        lset(&mut helper, names_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let assoc_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        ref_func(&mut helper, helper_idx, line);
        dynamic_get_from_slots(&mut helper, assoc_slot, key_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, out_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, assoc_loop, line);
        lget(&mut helper, names_slot, line);
        push_str(&mut helper, "\x1F", line);
        call_import_into(imports, &mut helper, "ecma:array", "join", 2, line);
        lset(&mut helper, tmp_slot, line);
        set_struct_from_slot(&mut helper, out_slot, "vybe$assoc_keys_csv", tmp_slot, line);
        helper.emit_end(line);
        helper.emit_end(line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);

        helper.emit_end(line);

        ref_func(&mut helper, alloc_idx, line);
        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "class", line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, out_slot, line);

        lget(&mut helper, kind_slot, line);
        push_str(&mut helper, "custom_object", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
        helper.emit_if(line);
        lget(&mut helper, out_slot, line);
        struct_get_key(&mut helper, "__unserialize", line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_uns = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_uns, line);
            lget(&mut helper, fn_slot_uns, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_uns, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_uns, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_uns, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        lget(&mut helper, method_slot, line);
        lget(&mut helper, out_slot, line);
        ref_func(&mut helper, helper_idx, line);
        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "payload", line);
        call_ref(&mut helper, 1, line);
        call_ref(&mut helper, 2, line);
        helper.emit_op(Op::DROP, line);
        helper.emit_end(line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);
        helper.emit_end(line);

        lget(&mut helper, node_slot, line);
        struct_get_key(&mut helper, "fields", line);
        lset(&mut helper, fields_slot, line);
        lget(&mut helper, fields_slot, line);
        call_import_into(imports, &mut helper, "ecma:object", "keys", 1, line);
        lset(&mut helper, names_slot, line);
        lget(&mut helper, names_slot, line);
        helper.emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut helper, n_slot, line);
        push_const(&mut helper, Value::F64(0.0), line);
        lset(&mut helper, i_slot, line);
        let fields_loop = helper_loop_start(&mut helper, line);
        lget(&mut helper, i_slot, line);
        lget(&mut helper, n_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut helper, line);
        helper_loop_cond(&mut helper, line);
        lget(&mut helper, names_slot, line);
        lget(&mut helper, i_slot, line);
        helper.emit_op(Op::ARRAY_GET, line);
        lset(&mut helper, key_slot, line);
        ref_func(&mut helper, helper_idx, line);
        dynamic_get_from_slots(&mut helper, fields_slot, key_slot, line);
        call_ref(&mut helper, 1, line);
        lset(&mut helper, tmp_slot, line);
        dynamic_set_from_slots(&mut helper, out_slot, key_slot, tmp_slot, line);
        bump_loop_index(&mut helper, i_slot, line);
        helper_loop_end(&mut helper, fields_loop, line);

        lget(&mut helper, kind_slot, line);
        push_str(&mut helper, "sleep_object", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut helper, line);
        helper.emit_if(line);
        lget(&mut helper, out_slot, line);
        struct_get_key(&mut helper, "__wakeup", line);
        lset(&mut helper, method_slot, line);
        // function test: not null AND not number AND not string AND not boolean
        {
            let fn_slot_wk = alloc_local(&mut helper);
            lget(&mut helper, method_slot, line);
            helper.emit_op_u16(Op::LOCAL_SET, fn_slot_wk, line);
            lget(&mut helper, fn_slot_wk, line);
            helper.emit_op(Op::REF_IS_NULL, line);
            helper.emit_op(Op::I32_EQZ, line);
            lget(&mut helper, fn_slot_wk, line);
            let tn = helper.add_import("wasm:js-number", "test");
            helper.emit_call(tn, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_wk, line);
            let ts = helper.add_import("wasm:js-string", "test");
            helper.emit_call(ts, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
            lget(&mut helper, fn_slot_wk, line);
            let tb = helper.add_import("wasm:js-boolean", "test");
            helper.emit_call(tb, 1, line);
            helper.emit_op(Op::I32_EQZ, line);
            helper.emit_op(Op::I32_AND, line);
        }
        helper.emit_if(line);
        lget(&mut helper, method_slot, line);
        lget(&mut helper, out_slot, line);
        call_ref(&mut helper, 1, line);
        helper.emit_op(Op::DROP, line);
        helper.emit_end(line);
        helper.emit_end(line);
        lget(&mut helper, out_slot, line);
        helper.emit_op(Op::RETURN, line);

        helper.emit_end(line);
        helper.emit_end(line);
    }

    chunks.push(helper);
    helper_idx
}

pub fn emit_php_empty(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    let _type_slot = alloc_local(chunk);

    lset(chunk, value_slot, line);

    lget(chunk, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    // boolean test
    lget(chunk, value_slot, line);
    let test_bool_empty = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_empty, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_else(line);

    // number test
    lget(chunk, value_slot, line);
    let test_num_empty = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(test_num_empty, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);

    // string test
    lget(chunk, value_slot, line);
    let test_str_empty = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_empty, 1, line);
    chunk.emit_if_value(line);

    lget(chunk, value_slot, line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    push_str(chunk, "0", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    // Arrays must be tested before object metadata: JS Array-backed PHP arrays
    // can expose internal fields, but PHP truthiness is length-based.
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    core_wasm::i32_const(chunk, line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    struct_get_key(chunk, "__type", line);
    let type_slot = alloc_local(chunk);
    lset(chunk, type_slot, line);
    lget(chunk, type_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    lget(chunk, type_slot, line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);

    // array test via ecma:array.isArray
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    let base_len_slot = alloc_local(chunk);
    let extra_len_slot = alloc_local(chunk);
    lset(chunk, base_len_slot, line);

    lget(chunk, value_slot, line);
    let assoc_key = chunk.add_constant(Value::String(Arc::from("vybe$assoc_keys_csv")));
    chunk.emit_op_u16(Op::STRUCT_GET, assoc_key, line);
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, base_len_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);
    push_str(chunk, "\x1F", line);
    let _ = chunk;
    call_import(chunks, current, "ecma:string", "split", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, extra_len_slot, line);
    lget(chunk, base_len_slot, line);
    lget(chunk, extra_len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    core_wasm::i32_const(chunk, line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_php_serialize(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let helper_idx = build_php_serialize_helper(chunks, line);
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    lset(chunk, value_slot, line);
    ref_func(chunk, helper_idx, line);
    lget(chunk, value_slot, line);
    call_ref(chunk, 1, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:json", "stringify", 1, line);
}

pub fn emit_php_unserialize(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let alloc_idx = build_php_alloc_helper(chunks, line);
    let helper_idx = build_php_unserialize_helper(chunks, alloc_idx, line);
    let chunk = &mut chunks[current];
    if argc > 1 {
        let _options_slot = alloc_local(chunk);
        lset(chunk, _options_slot, line);
    }
    let value_slot = alloc_local(chunk);
    lset(chunk, value_slot, line);
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:json", "parse", 1, line);
    let chunk = &mut chunks[current];
    let parsed_slot = alloc_local(chunk);
    lset(chunk, parsed_slot, line);
    ref_func(chunk, helper_idx, line);
    lget(chunk, parsed_slot, line);
    call_ref(chunk, 1, line);
}

/// `WeakReference::create($obj)` → struct with `get` method that derefs the weak ref.
pub fn emit_weak_ref_create(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    // Build the `get()` method: deref the weak ref stored in this.__weak
    let get_idx = {
        let mut c = Chunk::new("__weak_ref_get");
        c.arity = 1; // this
        let weak_k = c.add_constant(Value::String(Arc::from("__weak")));
        c.emit_op_u16(Op::LOCAL_GET, 0, line);
        c.emit_op_u16(Op::STRUCT_GET, weak_k, line);
        {
            let idx = c.add_import("ecma:weakref", "deref");
            c.emit_call(idx, 1, line);
        }
        c.emit_op(Op::RETURN, line);
        c.local_count = c.local_count.max(1);
        chunks.push(c);
        chunks.len() - 1
    };

    let chunk = &mut chunks[current];
    let obj_slot = chunk.alloc_scratch(1);
    let this_slot = chunk.alloc_scratch(1);

    // Pop the object arg
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    // Create struct, store weak ref
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);

    // this.__weak = REF_MAKE_WEAK(obj)
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    {
        let idx = chunk.add_import("ecma:weakref", "new");
        chunk.emit_call(idx, 1, line);
    }
    let weak_k = chunk.add_constant(Value::String(Arc::from("__weak")));
    chunk.emit_op_u16(Op::STRUCT_SET, weak_k, line);
    chunk.emit_op(Op::DROP, line);

    // Bind get method
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, get_idx as u16, line);
    chunk.emit(0, line);
    let get_k = chunk.add_constant(Value::String(Arc::from("get")));
    chunk.emit_op_u16(Op::STRUCT_SET, get_k, line);
    chunk.emit_op(Op::DROP, line);

    // Return the struct
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}


/// PHP `setcookie($name, $value, $options)`.
///
/// PHP URL-ENCODES the value here (`urlencode`, so `-_.` survive and a space
/// becomes `+`); `setrawcookie()` is the same call without that step. That
/// choice is PHP's, which is why it lives in this adapter and not in
/// `primitives/http_cookie.rs` — the primitive serializes verbatim.
pub fn emit_php_setcookie(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_setcookie_inner(chunks, current, argc, true, line);
}

/// PHP `setrawcookie()` — the value is sent exactly as given.
pub fn emit_php_setrawcookie(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_setcookie_inner(chunks, current, argc, false, line);
}

fn emit_setcookie_inner(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    encode_value: bool,
    line: u32,
) {
    let attrs_slot = if argc >= 3 {
        Some(alloc_local(&mut chunks[current]))
    } else {
        None
    };
    let value_slot = alloc_local(&mut chunks[current]);
    let name_slot = alloc_local(&mut chunks[current]);

    if let Some(slot) = attrs_slot {
        lset(&mut chunks[current], slot, line);
    }
    if argc >= 2 {
        lset(&mut chunks[current], value_slot, line);
    } else {
        push_str(&mut chunks[current], "", line);
        lset(&mut chunks[current], value_slot, line);
    }
    lset(&mut chunks[current], name_slot, line);

    lget(&mut chunks[current], name_slot, line);
    lget(&mut chunks[current], value_slot, line);
    if encode_value {
        crate::emitter::string_adapter::emit_urlencode(chunks, current, 1, line);
    }
    if let Some(slot) = attrs_slot {
        lget(&mut chunks[current], slot, line);
    }
    vybe_compiler::primitives::http_cookie::emit_serialize(
        chunks,
        current,
        if attrs_slot.is_some() { 3 } else { 2 },
        line,
    );

    // Reuse the ordinary header path — a cookie is just a `Set-Cookie` header,
    // which is why writing one needs no host function of its own.
    emit_send_cookie(chunks, current, line);
    push_const(&mut chunks[current], Value::Bool(true), line);
}
