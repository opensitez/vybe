use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::object::emit_bind_method_with_slot;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::{field_slot, set_both_spellings};

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
    let value_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, src, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    set_both_spellings(chunk, obj_slot, value_slot, dest, line);
}

fn emit_uri_port_value_from_obj(chunk: &mut Chunk, obj_slot: u16, line: u32) {
    let raw_port = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "port", line);
    host::emit(chunk, "ecma:number", "Number", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, raw_port, line);

    chunk.emit_op_u16(Op::LOCAL_GET, raw_port, line);
    chunk.emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "protocol", line);
    push_str(chunk, "https:", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(443, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "protocol", line);
    push_str(chunk, "http:", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(80, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw_port, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw_port, line);
    chunk.emit_end(line);
}

fn struct_set_drop_both(chunk: &mut Chunk, field: &str, line: u32) {
    let value_slot = reserve_slot(chunk);
    let obj_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    set_both_spellings(chunk, obj_slot, value_slot, field, line);
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

    set_alias_from_field(chunk, obj_slot, "host", "Authority", line);
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
    struct_set_drop_both(chunk, "Scheme", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "protocol", line);
    push_str(chunk, ":", line);
    push_str(chunk, "", line);
    host::emit(chunk, "ecma:string", "replace", 3, line);
    struct_set_drop(chunk, "scheme", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    emit_uri_port_value_from_obj(chunk, obj_slot, line);
    struct_set_drop_both(chunk, "Port", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "protocol", line);
    push_str(chunk, "file:", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    struct_set_drop_both(chunk, "IsFile", line);

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
    struct_set_drop_both(chunk, "HostNameType", line);

    // The rest of `System.Uri`'s read surface, every member a pure function of
    // what WHATWG already parsed. Values checked against pwsh 7.6.4 on
    // `https://user:pw@example.com:8443/a/b?q=1&p=2#frag`.
    set_alias_from_field(chunk, obj_slot, "hostname", "DnsSafeHost", line);
    set_alias_from_field(chunk, obj_slot, "pathname", "LocalPath", line);
    set_alias_from_field(chunk, obj_slot, "href", "AbsoluteUri", line);
    set_alias_from_field(chunk, obj_slot, "href", "OriginalString", line);
    set_alias_from_field(chunk, obj_slot, "href", "__value", line);

    // `Uri` is what `GetType().Name` owes, whatever spelling reached the cast.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_str(chunk, "Uri", line);
    struct_set_drop_both(chunk, "__type", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_str(chunk, "Uri", line);
    struct_set_drop_both(chunk, "name", line);

    // A parsed absolute URL is absolute by construction; the relative form
    // never reaches this finalizer.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    vybe_compiler::primitives::instructions::core_wasm::bool_const(chunk, line, true);
    struct_set_drop_both(chunk, "IsAbsoluteUri", line);

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
    struct_set_drop_both(chunk, "UserInfo", line);

    // `PathAndQuery` keeps the `?`, which is how `search` already spells it.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "pathname", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "search", line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    struct_set_drop_both(chunk, "PathAndQuery", line);

    // Loopback is the host, not the address family: .NET answers True for
    // `127.0.0.1`, `localhost` and `::1`.
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_str(chunk, "^(127\\.|localhost$|\\[?::1\\]?$)", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    struct_get(chunk, "hostname", line);
    host::emit(chunk, "ecma:regexp", "test", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
    struct_set_drop_both(chunk, "IsLoopback", line);

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
    struct_set_drop_both(chunk, "Segments", line);

    bind_uri_methods(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

fn emit_relative_uri_from_slot(chunks: &mut Vec<Chunk>, current: usize, text_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    class_slots::emit_class_alloc(chunk, line);
    let uri_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, uri_slot, line);

    for field in ["href", "OriginalString", "__value", "PathAndQuery"] {
        set_both_spellings(chunk, uri_slot, text_slot, field, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, uri_slot, line);
    push_str(chunk, "Uri", line);
    struct_set_drop_both(chunk, "__type", line);
    chunk.emit_op_u16(Op::LOCAL_GET, uri_slot, line);
    push_str(chunk, "Uri", line);
    struct_set_drop_both(chunk, "name", line);
    chunk.emit_op_u16(Op::LOCAL_GET, uri_slot, line);
    vybe_compiler::primitives::instructions::core_wasm::bool_const(chunk, line, false);
    struct_set_drop_both(chunk, "IsAbsoluteUri", line);

    bind_uri_tostring_only(chunks, current, uri_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, uri_slot, line);
}

fn set_builder_defaults(chunk: &mut Chunk, builder_slot: u16, line: u32) {
    let value = reserve_slot(chunk);
    for (field, text) in [
        ("Scheme", "http"),
        ("Host", "localhost"),
        ("Path", "/"),
        ("Query", ""),
        ("Fragment", ""),
        ("UserName", ""),
        ("Password", ""),
    ] {
        push_str(chunk, text, line);
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
        set_both_spellings(chunk, builder_slot, value, field, line);
    }
    chunk.emit_i32_const(-1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    set_both_spellings(chunk, builder_slot, value, "Port", line);
    push_str(chunk, "UriBuilder", line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    set_both_spellings(chunk, builder_slot, value, "__type", line);
    set_both_spellings(chunk, builder_slot, value, "name", line);
}

fn copy_uri_field_to_builder(
    chunk: &mut Chunk,
    builder_slot: u16,
    uri_slot: u16,
    uri_field: &str,
    builder_field: &str,
    line: u32,
) {
    let value = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, uri_slot, line);
    struct_get(chunk, uri_field, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    set_both_spellings(chunk, builder_slot, value, builder_field, line);
}

fn emit_uri_builder_from_uri_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    uri_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    class_slots::emit_class_alloc(chunk, line);
    let builder_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, builder_slot, line);
    set_builder_defaults(chunk, builder_slot, line);
    copy_uri_field_to_builder(chunk, builder_slot, uri_slot, "Scheme", "Scheme", line);
    copy_uri_field_to_builder(chunk, builder_slot, uri_slot, "Host", "Host", line);
    copy_uri_field_to_builder(chunk, builder_slot, uri_slot, "Port", "Port", line);
    copy_uri_field_to_builder(chunk, builder_slot, uri_slot, "AbsolutePath", "Path", line);
    copy_uri_field_to_builder(chunk, builder_slot, uri_slot, "Query", "Query", line);
    copy_uri_field_to_builder(chunk, builder_slot, uri_slot, "Fragment", "Fragment", line);
    chunk.emit_op_u16(Op::LOCAL_GET, builder_slot, line);
}

fn emit_builder_field(chunk: &mut Chunk, builder_slot: u16, field: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, builder_slot, line);
    struct_get(chunk, field, line);
}

fn emit_prefixed_part(
    chunk: &mut Chunk,
    value_slot: u16,
    prefix: &str,
    empty_value: &str,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, empty_value, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    push_str(chunk, prefix, line);
    host::emit(chunk, "ecma:string", "startsWith", 2, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_else(line);
    push_str(chunk, prefix, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_port_part(chunk: &mut Chunk, port_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, port_slot, line);
    chunk.emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, ":", line);
    chunk.emit_op_u16(Op::LOCAL_GET, port_slot, line);
    vybe_compiler::primitives::strings::emit_to_string(chunk, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    chunk.emit_else(line);
    push_str(chunk, "", line);
    chunk.emit_end(line);
}

fn emit_userinfo_part(chunk: &mut Chunk, username_slot: u16, password_slot: u16, line: u32) {
    let password_part = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, username_slot, line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, password_slot, line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "", line);
    chunk.emit_else(line);
    push_str(chunk, ":", line);
    chunk.emit_op_u16(Op::LOCAL_GET, password_slot, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, password_part, line);
    chunk.emit_op_u16(Op::LOCAL_GET, username_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, password_part, line);
    push_str(chunk, "@", line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 3, line);
    chunk.emit_end(line);
}

fn append_field_part(chunk: &mut Chunk, href_slot: u16, part_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, href_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, part_slot, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, href_slot, line);
}

pub fn emit_uri_builder_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let builder = reserve_slot(chunk);
    let href = reserve_slot(chunk);
    let part = reserve_slot(chunk);
    let port = reserve_slot(chunk);
    let username = reserve_slot(chunk);
    let password = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, builder, line);

    emit_builder_field(chunk, builder, "scheme", line);
    push_str(chunk, "://", line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, href, line);

    emit_builder_field(chunk, builder, "username", line);
    chunk.emit_op_u16(Op::LOCAL_SET, username, line);
    emit_builder_field(chunk, builder, "password", line);
    chunk.emit_op_u16(Op::LOCAL_SET, password, line);
    emit_userinfo_part(chunk, username, password, line);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line);
    append_field_part(chunk, href, part, line);

    emit_builder_field(chunk, builder, "host", line);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line);
    append_field_part(chunk, href, part, line);

    emit_builder_field(chunk, builder, "port", line);
    chunk.emit_op_u16(Op::LOCAL_SET, port, line);
    emit_port_part(chunk, port, line);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line);
    append_field_part(chunk, href, part, line);

    emit_builder_field(chunk, builder, "path", line);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line);
    emit_prefixed_part(chunk, part, "/", "/", line);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line);
    append_field_part(chunk, href, part, line);

    emit_builder_field(chunk, builder, "query", line);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line);
    emit_prefixed_part(chunk, part, "?", "", line);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line);
    append_field_part(chunk, href, part, line);

    emit_builder_field(chunk, builder, "fragment", line);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line);
    emit_prefixed_part(chunk, part, "#", "", line);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line);
    append_field_part(chunk, href, part, line);

    chunk.emit_op_u16(Op::LOCAL_GET, href, line);
}

pub fn emit_uri_builder_uri(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_uri_builder_to_string(chunks, current, line);
    emit_uri_new(chunks, current, 1, line);
}

pub fn emit_uri_builder_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    match argc {
        0 => {
            let chunk = &mut chunks[current];
            class_slots::emit_class_alloc(chunk, line);
            let builder = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, builder, line);
            set_builder_defaults(chunk, builder, line);
            chunk.emit_op_u16(Op::LOCAL_GET, builder, line);
        }
        _ => {
            let arg_slot;
            {
                let chunk = &mut chunks[current];
                for _ in 1..argc {
                    chunk.emit_op(Op::DROP, line);
                }
                arg_slot = reserve_slot(chunk);
                chunk.emit_op_u16(Op::LOCAL_SET, arg_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
                struct_get(chunk, "__type", line);
                push_str(chunk, "Uri", line);
                vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
                chunk.emit_if_value(line);
            }
            emit_uri_builder_from_uri_slot(chunks, current, arg_slot, line);
            {
                let chunk = &mut chunks[current];
                chunk.emit_else(line);
                chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
            }
            emit_uri_new(chunks, current, 1, line);
            let parsed_slot;
            {
                let chunk = &mut chunks[current];
                parsed_slot = reserve_slot(chunk);
                chunk.emit_op_u16(Op::LOCAL_SET, parsed_slot, line);
            }
            emit_uri_builder_from_uri_slot(chunks, current, parsed_slot, line);
            chunks[current].emit_end(line);
        }
    }
}

pub fn emit_uri_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let url_idx = chunks[current].add_import("node:url", "URL");
    match argc {
        2 => {
            let relative_slot;
            let base_slot;
            {
                let chunk = &mut chunks[current];
                relative_slot = reserve_slot(chunk);
                base_slot = reserve_slot(chunk);
                chunk.emit_op_u16(Op::LOCAL_SET, relative_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, base_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, relative_slot, line);
                push_str(chunk, "Absolute", line);
                vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
                chunk.emit_op_u16(Op::LOCAL_GET, relative_slot, line);
                push_str(chunk, "Relative", line);
                vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
                chunk.emit_op(Op::I32_OR, line);
                chunk.emit_op_u16(Op::LOCAL_GET, relative_slot, line);
                push_str(chunk, "RelativeOrAbsolute", line);
                vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
                chunk.emit_op(Op::I32_OR, line);
                chunk.emit_if_value(line);
                chunk.emit_op_u16(Op::LOCAL_GET, relative_slot, line);
                push_str(chunk, "Relative", line);
                vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
                chunk.emit_if_value(line);
            }
            emit_relative_uri_from_slot(chunks, current, base_slot, line);
            {
                let chunk = &mut chunks[current];
                chunk.emit_else(line);
                chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
                chunk.emit_call(url_idx, 1, line);
            }
            emit_finalize_uri(chunks, current, line);
            {
                let chunk = &mut chunks[current];
                chunk.emit_end(line);
                chunk.emit_else(line);
                chunk.emit_op_u16(Op::LOCAL_GET, relative_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
                struct_get(chunk, "href", line);
                chunk.emit_call(url_idx, 2, line);
            }
            emit_finalize_uri(chunks, current, line);
            chunks[current].emit_end(line);
        }
        _ => {
            let chunk = &mut chunks[current];
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            let input_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
            host::emit(chunk, "node:url", "canParse", 1, line);
            chunk.emit_op(Op::I32_EQZ, line);
            chunk.emit_if(line);
            crate::emitter::core::exceptions::emit_new_typed(
                chunks,
                current,
                "UriFormatException",
                class_slots::ValueSource::ConstStr(
                    "Invalid URI: The format of the URI could not be determined.".to_string(),
                ),
                line,
            );
            let chunk = &mut chunks[current];
            vybe_compiler::primitives::errors::emit_throw(chunk, line);
            chunk.emit_end(line);
            chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
            chunk.emit_call(url_idx, 1, line);
            emit_finalize_uri(chunks, current, line);
        }
    }
}

pub fn emit_uri_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get(&mut chunks[current], "href", line);
}

pub fn emit_uri_port(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_uri_port_value_from_obj(chunk, obj_slot, line);
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
