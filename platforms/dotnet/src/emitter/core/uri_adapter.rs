use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::object::emit_bind_method_with_slot;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn struct_get(chunk: &mut Chunk, field: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(field),
        Dest::Stack,
        line,
    );
}

fn struct_set_drop(chunk: &mut Chunk, field: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(field),
        ValueSource::Stack,
        line,
    );
}

fn set_alias_from_field(chunk: &mut Chunk, obj_slot: u16, src: &str, dest: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, src, line);
    struct_set_drop(chunk, dest, line);
}

fn bind_uri_methods(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let mut tostring = create_function_chunk("__dotnet_uri_tostring", 1);
    tostring.emit_op_u16(Op::LOCAL_GET, 0, line);
    class_slots::emit_class_get(
        &mut tostring,
        ObjSource::Stack,
        &field_slot("href"),
        Dest::Stack,
        line,
    );
    tostring.emit_op(Op::RETURN, line);
    tostring.local_count = 1;
    chunks.push(tostring);
    let tostring_idx = chunks.len() - 1;

    let mut is_base_of = create_function_chunk("__dotnet_uri_isbaseof", 2);
    let starts_with_idx = is_base_of.add_import("ecma:string", "startsWith");
    is_base_of.emit_op_u16(Op::LOCAL_GET, 1, line);
    class_slots::emit_class_get(
        &mut is_base_of,
        ObjSource::Stack,
        &field_slot("href"),
        Dest::Stack,
        line,
    );
    is_base_of.emit_op_u16(Op::LOCAL_GET, 0, line);
    class_slots::emit_class_get(
        &mut is_base_of,
        ObjSource::Stack,
        &field_slot("href"),
        Dest::Stack,
        line,
    );
    is_base_of.emit_call(starts_with_idx, 2, line);
    is_base_of.emit_op(Op::RETURN, line);
    is_base_of.local_count = 2;
    chunks.push(is_base_of);
    let is_base_of_idx = chunks.len() - 1;

    let mut make_relative = create_function_chunk("__dotnet_uri_makerelative", 2);
    let replace_idx = make_relative.add_import("ecma:string", "replace");
    let relative_slot = make_relative.alloc_scratch(2);
    let uri_slot = relative_slot + 1;
    make_relative.emit_op_u16(Op::LOCAL_GET, 1, line);
    class_slots::emit_class_get(
        &mut make_relative,
        ObjSource::Stack,
        &field_slot("href"),
        Dest::Stack,
        line,
    );
    make_relative.emit_op_u16(Op::LOCAL_GET, 0, line);
    class_slots::emit_class_get(
        &mut make_relative,
        ObjSource::Stack,
        &field_slot("href"),
        Dest::Stack,
        line,
    );
    push_str(&mut make_relative, "", line);
    make_relative.emit_call(replace_idx, 3, line);
    make_relative.emit_op_u16(Op::LOCAL_SET, relative_slot, line);
    class_slots::emit_class_alloc(&mut make_relative, line);
    make_relative.emit_op_u16(Op::LOCAL_SET, uri_slot, line);
    make_relative.emit_op_u16(Op::LOCAL_GET, uri_slot, line);
    make_relative.emit_op_u16(Op::LOCAL_GET, relative_slot, line);
    class_slots::emit_class_set(
        &mut make_relative,
        ObjSource::Stack,
        &field_slot("href"),
        ValueSource::Stack,
        line,
    );
    emit_bind_method_with_slot(
        &mut make_relative,
        uri_slot,
        "tostring",
        Some(vybe_ast::ProtocolSlot::ToString),
        tostring_idx,
        None,
        line,
    );
    make_relative.emit_op_u16(Op::LOCAL_GET, uri_slot, line);
    make_relative.emit_op(Op::RETURN, line);
    make_relative.local_count = 2;
    chunks.push(make_relative);
    let make_relative_idx = chunks.len() - 1;

    let chunk = &mut chunks[current];
    emit_bind_method_with_slot(
        chunk,
        obj_slot,
        "tostring",
        Some(vybe_ast::ProtocolSlot::ToString),
        tostring_idx,
        None,
        line,
    );
    emit_bind_method_with_slot(
        chunk,
        obj_slot,
        "isbaseof",
        None,
        is_base_of_idx,
        None,
        line,
    );
    emit_bind_method_with_slot(
        chunk,
        obj_slot,
        "IsBaseOf",
        None,
        is_base_of_idx,
        None,
        line,
    );
    emit_bind_method_with_slot(
        chunk,
        obj_slot,
        "makerelativeuri",
        None,
        make_relative_idx,
        None,
        line,
    );
    emit_bind_method_with_slot(
        chunk,
        obj_slot,
        "MakeRelativeUri",
        None,
        make_relative_idx,
        None,
        line,
    );
}

fn emit_finalize_uri(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    set_alias_from_field(chunk, obj_slot, "hostname", "Host", line);
    set_alias_from_field(chunk, obj_slot, "hostname", "host", line);
    set_alias_from_field(chunk, obj_slot, "pathname", "AbsolutePath", line);
    set_alias_from_field(chunk, obj_slot, "search", "Query", line);
    set_alias_from_field(chunk, obj_slot, "hash", "Fragment", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "protocol", line);
    push_str(chunk, ":", line);
    push_str(chunk, "", line);
    host::emit(chunk, "ecma:string", "replace", 3, line);
    struct_set_drop(chunk, "Scheme", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "port", line);
    host::emit(chunk, "ecma:number", "parseInt", 1, line);
    struct_set_drop(chunk, "Port", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "protocol", line);
    push_str(chunk, "file:", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    struct_set_drop(chunk, "IsFile", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_str(chunk, "^\\d+\\.", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "hostname", line);
    host::emit(chunk, "ecma:regexp", "test", 2, line);
    chunk.emit_if_value(line);
    push_str(chunk, "IPv4", line);
    chunk.emit_else(line);
    push_str(chunk, "Dns", line);
    chunk.emit_end(line);
    struct_set_drop(chunk, "HostNameType", line);

    // The rest of `System.Uri`'s read surface, every member a pure function of
    // what WHATWG already parsed. Values checked against pwsh 7.6.4 on
    // `https://user:pw@example.com:8443/a/b?q=1&p=2#frag`.
    set_alias_from_field(chunk, obj_slot, "hostname", "DnsSafeHost", line);
    set_alias_from_field(chunk, obj_slot, "pathname", "LocalPath", line);
    set_alias_from_field(chunk, obj_slot, "href", "AbsoluteUri", line);
    set_alias_from_field(chunk, obj_slot, "href", "OriginalString", line);
    set_alias_from_field(chunk, obj_slot, "host", "Authority", line);

    // `Uri` is what `GetType().Name` owes, whatever spelling reached the cast.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_str(chunk, "Uri", line);
    struct_set_drop(chunk, "__type", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_str(chunk, "Uri", line);
    struct_set_drop(chunk, "name", line);

    // A parsed absolute URL is absolute by construction; the relative form
    // never reaches this finalizer.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    vybe_compiler::primitives::instructions::core_wasm::bool_const(chunk, line, true);
    struct_set_drop(chunk, "IsAbsoluteUri", line);

    // `UserInfo` is `user[:pass]`, and EMPTY when there is no user — not
    // `":"`, which a bare concatenation would give.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "username", line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "username", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "password", line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    chunk.emit_else(line);
    push_str(chunk, ":", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "password", line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    chunk.emit_end(line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    chunk.emit_end(line);
    struct_set_drop(chunk, "UserInfo", line);

    // `PathAndQuery` keeps the `?`, which is how `search` already spells it.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "pathname", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "search", line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    struct_set_drop(chunk, "PathAndQuery", line);

    // Loopback is the host, not the address family: .NET answers True for
    // `127.0.0.1`, `localhost` and `::1`.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_str(chunk, "^(127\\.|localhost$|\\[?::1\\]?$)", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "hostname", line);
    host::emit(chunk, "ecma:regexp", "test", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    struct_set_drop(chunk, "IsLoopback", line);

    // `Segments` keeps each separator: `/a/b/c/` is `/`, `a/`, `b/`, `c/`, and
    // `/a/b` is `/`, `a/`, `b`. One global match expresses both — a run up to
    // and including a slash, or a final run without one.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "pathname", line);
    // ⛔THE FLAGS RIDE IN THE PATTERN. `extract_pattern` reads `/pat/flags`
    // out of the string it is handed; a third argument is not read, so a
    // separate `"g"` left the match non-global and `Segments` answered `/`
    // alone. The inner slashes are escaped because the delimiter is a slash.
    push_str(chunk, "/[^\\/]*\\/|[^\\/]+$/g", line);
    // `match` with the global flag answers an array of the matched STRINGS,
    // which is the shape `Segments` is; `matchAll` answers match objects.
    host::emit(chunk, "ecma:regexp", "match", 2, line);
    struct_set_drop(chunk, "Segments", line);

    bind_uri_methods(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_uri_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let url_idx = chunks[current].add_import("node:url", "URL");
    let chunk = &mut chunks[current];
    match argc {
        2 => {
            let relative_slot = reserve_slot(chunk);
            let base_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, relative_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, base_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, relative_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
            struct_get(chunk, "href", line);
            chunk.emit_call(url_idx, 2, line);
        }
        _ => {
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            let input_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
            host::emit(chunk, "node:url", "canParse", 1, line);
            chunk.emit_op(Op::I32_EQZ, line);
            chunk.emit_if(line);
            class_slots::emit_class_alloc(chunk, line);
            chunk.emit_dup(line);
            chunk.emit_string_const(
                "Invalid URI: The format of the URI could not be determined.",
                line,
            );
            vybe_compiler::primitives::errors::emit_exception_new_finalize(
                chunk,
                "UriFormatException",
                line,
            );
            vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
                chunk,
                "UriFormatException",
                line,
            );
            vybe_compiler::primitives::errors::emit_throw(chunk, line);
            chunk.emit_end(line);
            chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
            chunk.emit_call(url_idx, 1, line);
        }
    }
    emit_finalize_uri(chunks, current, line);
}

pub fn emit_uri_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get(&mut chunks[current], "href", line);
}

/// `Uri.EscapeDataString` — RFC 3986 percent-encoding.
///
/// Routed through the SHARED codec `primitives::url`, the same one php
/// `rawurlencode`, python `quote`, java `URLEncoder` and go `QueryEscape` use.
/// The four differ only by `PercentOptions`; .NET is the plain RFC 3986
/// variant. Behaviour is unchanged — `rfc3986()` emits exactly the
/// `encodeURIComponent` this used to call directly — but .NET now moves with
/// the codec instead of drifting from it.
///
/// Known pre-existing gap, NOT introduced here: real `Uri.EscapeDataString`
/// escapes `!*'()` (its unreserved set is only `A-Za-z0-9-._~`) where
/// `encodeURIComponent` leaves them. That is a future `PercentOptions` field,
/// not a reason to fork.
pub fn emit_uri_escape(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::url::emit_percent_encode(
        chunks,
        current,
        vybe_compiler::primitives::url::PercentOptions::rfc3986(),
        line,
    );
}

/// `Uri.UnescapeDataString` — the inverse, through the same shared codec.
/// `+` stays literal in RFC 3986 mode, which is .NET's behaviour.
pub fn emit_uri_unescape(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::url::emit_percent_decode(
        chunks,
        current,
        vybe_compiler::primitives::url::PercentOptions::rfc3986(),
        line,
    );
}

pub fn emit_uri_is_well_formed(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    host::emit(chunk, "node:url", "canParse", 1, line);
}

pub fn emit_uri_try_create(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    match argc {
        2 => {
            let kind_slot = reserve_slot(chunk);
            let input_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, kind_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
            host::emit(chunk, "node:url", "canParse", 1, line);
            chunk.emit_if_value(line);
            chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
            chunk.emit_op(Op::DROP, line);
            let url_idx = chunk.add_import("node:url", "URL");
            chunk.emit_call(url_idx, 1, line);
            emit_finalize_uri(chunks, current, line);
            chunks[current].emit_else(line);
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            chunks[current].emit_end(line);
        }
        _ => {
            if argc > 1 {
                for _ in 1..argc {
                    chunk.emit_op(Op::DROP, line);
                }
            }
            host::emit(chunk, "node:url", "canParse", 1, line);
        }
    }
}

pub fn emit_uri_kind(name: &str, chunks: &mut [Chunk], current: usize, line: u32) {
    push_str(&mut chunks[current], name, line);
}

pub fn emit_uri_is_base_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let child_slot = reserve_slot(chunk);
    let base_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, child_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, child_slot, line);
    struct_get(chunk, "href", line);
    chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
    struct_get(chunk, "href", line);
    host::emit(chunk, "ecma:string", "startsWith", 2, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

fn bind_uri_tostring_only(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let mut tostring = create_function_chunk("__dotnet_uri_relative_tostring", 1);
    tostring.emit_op_u16(Op::LOCAL_GET, 0, line);
    class_slots::emit_class_get(
        &mut tostring,
        ObjSource::Stack,
        &field_slot("href"),
        Dest::Stack,
        line,
    );
    tostring.emit_op(Op::RETURN, line);
    tostring.local_count = 1;
    chunks.push(tostring);
    let tostring_idx = chunks.len() - 1;
    emit_bind_method_with_slot(
        &mut chunks[current],
        obj_slot,
        "tostring",
        Some(vybe_ast::ProtocolSlot::ToString),
        tostring_idx,
        None,
        line,
    );
}

pub fn emit_uri_make_relative(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let target_slot = reserve_slot(chunk);
    let base_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, target_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    struct_get(chunk, "href", line);
    chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
    struct_get(chunk, "href", line);
    push_str(chunk, "", line);
    host::emit(chunk, "ecma:string", "replace", 3, line);
    let relative_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, relative_slot, line);
    class_slots::emit_class_alloc(chunk, line);
    let uri_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, uri_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, uri_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, relative_slot, line);
    struct_set_drop(chunk, "href", line);
    bind_uri_tostring_only(chunks, current, uri_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, uri_slot, line);
}
