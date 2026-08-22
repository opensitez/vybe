use std::sync::Arc;
use vybe_compiler::primitives::{fs_path, paths};
use vybe_compiler::primitives::instructions::{core_wasm, host};

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::{collections, loops};

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn set_field_from_slot(chunk: &mut Chunk, obj_slot: u16, name: &str, value_slot: u16, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

fn set_field_const(chunk: &mut Chunk, obj_slot: u16, name: &str, val: Value, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, val, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

fn set_field_with_stack_value(chunk: &mut Chunk, obj_slot: u16, name: &str, line: u32) {
    let value_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    set_field_from_slot(chunk, obj_slot, name, value_slot, line);
}

// ── The CLR's failure model, laid over WASI's ─────────────────────────────
//
// `fs_path`'s verbs answer `null` or `false` on failure, and they are right to:
// WASI returns `result<_, error-code>` and each language maps that to its own
// idiom — PHP's `false`, Python's raise, Node's throw. **.NET's idiom is an
// exception**, and it is the whole reason these adapters exist. A `File.
// ReadAllText` that hands back `null` is not a .NET API; the caller's next
// `.Length` fails somewhere else entirely, or worse, `"" + null` succeeds and
// the missing file becomes an empty string.
//
// So the boundary is here, not in `fs_path`: WASI stays honest about results,
// and the .NET surface converts them to the exceptions `System.IO` documents.
// Verified against real .NET with `tools/csrun` — types AND message text, since
// a caller may print `e.Message`.

/// Throw a `System.IO` exception naming `path_slot`. Stack: `[]` → diverges.
///
/// The message is .NET's, verbatim: reads say `Could not find file '<path>'.`
/// and writes say `Could not find a part of the path '<path>'.`, both with the
/// FULL path — `File.ReadAllText("missing.txt")` reports the resolved absolute
/// path, not the relative string it was given.
fn emit_throw_io(
    chunks: &mut [Chunk],
    current: usize,
    kind: &str,
    prefix: &str,
    path_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const(prefix, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    paths::emit_full_path(chunk, line);
    chunk.emit_string_const("'.", line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 3, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(chunk, kind, line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

/// `FileNotFoundException` when the value on the stack is null.
/// Stack: `[value]` → `[value]`, or diverges.
///
/// `REF_IS_NULL` reports `Undefined` as null too, which is what makes this a
/// total test rather than a guess about which flavour of absent arrived.
fn emit_read_or_throw(chunks: &mut [Chunk], current: usize, path_slot: u16, line: u32) {
    let out = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_throw_io(
        chunks,
        current,
        "FileNotFoundException",
        "Could not find file '",
        path_slot,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `DirectoryNotFoundException` when the write reported failure.
/// Stack: `[bool]` → `[]`, or diverges.
///
/// .NET's write verbs return `void`, so the success bool is CONSUMED here
/// rather than pushed on — dropping it at the call site instead is how a failed
/// write became a silent success.
fn emit_write_or_throw(chunks: &mut [Chunk], current: usize, path_slot: u16, line: u32) {
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_throw_io(
        chunks,
        current,
        "DirectoryNotFoundException",
        "Could not find a part of the path '",
        path_slot,
        line,
    );
    chunks[current].emit_end(line);
}

/// `Directory.GetCurrentDirectory()` / `Environment.CurrentDirectory`.
/// Stack: `[]` → `[string]`.
///
/// `wasi:cli/environment.get-initial-cwd`, not `node:process.cwd`. INITIAL is
/// not an approximation of CURRENT here: WASI has no `chdir`, so a component's
/// starting directory is the only one it ever has. `node:process.cwd` reported
/// the HOST process's directory, which is a different process's state — it
/// happened to agree whenever nothing had tried to change it.
pub fn emit_current_directory(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let get_cwd = chunk.add_import("wasi:cli/environment", "get-initial-cwd");
    let cwd = reserve_slot(chunk);
    chunk.emit_call(get_cwd, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cwd, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cwd, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    // `""`, not null: the .NET property is typed `string` and every caller
    // concatenates or `Path.Combine`s it.
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, cwd, line);
    chunk.emit_end(line);
}

/// `File.ReadAllText(path)` — WASI underneath, `FileNotFoundException` on top.
/// Stack: `[path]` → `[string]`.
pub fn emit_file_read_all_text(chunks: &mut [Chunk], current: usize, line: u32) {
    let path = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_read_file(&mut chunks[current], line);
    emit_read_or_throw(chunks, current, path, line);
}

/// `File.WriteAllText(path, text)` / `AppendAllText`. Stack: `[path, text]` → `[]`.
pub fn emit_file_write_all_text(chunks: &mut [Chunk], current: usize, append: bool, line: u32) {
    let text = reserve_slot(&mut chunks[current]);
    let path = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    if append {
        fs_path::emit_append_file(&mut chunks[current], line);
    } else {
        fs_path::emit_write_file(&mut chunks[current], line);
    }
    emit_write_or_throw(chunks, current, path, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_file_stream_object(
    chunks: &mut [Chunk],
    current: usize,
    path_slot: u16,
    content_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);

    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    set_field_const(
        chunk,
        obj_slot,
        "__type",
        Value::String(Arc::from("FileStream")),
        line,
    );
    set_field_from_slot(chunk, obj_slot, "__path", path_slot, line);
    set_field_from_slot(chunk, obj_slot, "__content", content_slot, line);
    set_field_from_slot(chunk, obj_slot, "__buf", content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    set_field_with_stack_value(chunk, obj_slot, "Length", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `File.WriteAllBytes(path, bytes)`. Stack: `[path, bytes]` → `[null]`.
///
/// ⚠The inverse of [`emit_file_read_all_bytes`]'s bug, and worse because it
/// corrupted the FILE rather than the read: each byte became a character via
/// `String.fromCharCode`, the characters were concatenated, and the string was
/// written as UTF-8 — so every byte above 0x7F was written as TWO bytes and the
/// file was longer than the array. `io::Payload::Bytes` copies the array
/// verbatim.
pub fn emit_file_write_all_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let bytes = reserve_slot(&mut chunks[current]);
    let path = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bytes, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes, line);
    fs_path::emit_write_file_bytes(&mut chunks[current], line);
    emit_write_or_throw(chunks, current, path, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `File.ReadAllBytes(path)` → `byte[]`. Stack: `[path]` → `[array]`.
///
/// ⚠This read the file as UTF-8 TEXT and then took `charCodeAt` of each
/// character, which is not a byte read: any byte above 0x7F decoded into one
/// code point ≥ 0x80 (or U+FFFD for an invalid sequence), so the array came
/// back shorter than the file and with values the file never contained. It
/// happened to be right for pure ASCII, which is what every test held.
///
/// `read-via-stream` hands over the actual bytes, so there is nothing to
/// reconstruct.
pub fn emit_file_read_all_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let path = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_read_file_bytes(&mut chunks[current], line);
    emit_read_or_throw(chunks, current, path, line);
}

pub fn emit_file_create(chunks: &mut [Chunk], current: usize, line: u32) {
    let path_slot = reserve_slot(&mut chunks[current]);
    let content_slot = reserve_slot(&mut chunks[current]);
    let obj_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_op_u16(Op::LOCAL_SET, path_slot, line);
    push_const(&mut chunks[current], Value::String(Arc::from("")), line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, content_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    // `File.Create` TRUNCATES an existing file, which is `write-via-stream`
    // with the truncate open-flag — the same thing `writeFileSync` did.
    fs_path::emit_write_file(&mut chunks[current], line);
    emit_write_or_throw(chunks, current, path_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    set_field_const(
        chunk,
        obj_slot,
        "__type",
        Value::String(Arc::from("FileStream")),
        line,
    );
    set_field_from_slot(chunk, obj_slot, "__path", path_slot, line);
    set_field_from_slot(chunk, obj_slot, "__content", content_slot, line);
    set_field_from_slot(chunk, obj_slot, "__buf", content_slot, line);
    set_field_const(chunk, obj_slot, "Length", Value::I32(0), line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_file_open_read(chunks: &mut [Chunk], current: usize, line: u32) {
    let path_slot = reserve_slot(&mut chunks[current]);
    let content_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    fs_path::emit_read_file(&mut chunks[current], line);
    // `File.OpenRead` on a missing file throws before the stream object is
    // built — a `FileStream` wrapping `null` would fail on first read instead,
    // several frames from the call the user wrote.
    emit_read_or_throw(chunks, current, path_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, content_slot, line);
    emit_file_stream_object(chunks, current, path_slot, content_slot, line);
}

pub fn emit_file_stream_write_byte(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buf_key = chunk.add_constant(Value::String(Arc::from("__buf")));
    let path_key = chunk.add_constant(Value::String(Arc::from("__path")));
    let length_key = chunk.add_constant(Value::String(Arc::from("Length")));
    let stream_slot = reserve_slot(chunk);
    let byte_slot = reserve_slot(chunk);
    let buf_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, byte_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, stream_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, byte_slot, line);
    host::emit(chunk, "wasm:js-string", "fromCharCode", 1, line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, buf_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, length_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, path_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    fs_path::emit_write_file(chunk, line);
    // The write's success is DROPPED here and only here: the whole buffer is
    // rewritten on every `WriteByte`, so a failure surfaces on the next one or
    // on `Flush`, and throwing per byte would report the same failure once per
    // byte written.
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_file_info_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);
    let obj_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    set_field_const(
        chunk,
        obj_slot,
        "__type",
        Value::String(Arc::from("FileInfo")),
        line,
    );
    set_field_from_slot(chunk, obj_slot, "FullName", path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    paths::emit_file_name(chunk, line);
    set_field_with_stack_value(chunk, obj_slot, "Name", line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    paths::emit_extension(chunk, line);
    set_field_with_stack_value(chunk, obj_slot, "Extension", line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    fs_path::emit_exists(chunk, line);
    set_field_with_stack_value(chunk, obj_slot, "Exists", line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    fs_path::emit_file_size(chunk, line);
    set_field_with_stack_value(chunk, obj_slot, "Length", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_file_read_all_lines(chunks: &mut [Chunk], current: usize, line: u32) {
    let path_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    fs_path::emit_read_file(&mut chunks[current], line);
    // Thrown BEFORE the split: splitting `null` on "\n" yields an array with
    // one empty string, which is indistinguishable from an empty file.
    emit_read_or_throw(chunks, current, path_slot, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::String(Arc::from("\n")), line);
    host::emit(chunk, "ecma:string", "split", 2, line);
}

pub fn emit_file_write_all_lines(chunks: &mut [Chunk], current: usize, line: u32) {
    let path_slot = reserve_slot(&mut chunks[current]);
    let lines_slot = reserve_slot(&mut chunks[current]);
    let text_slot = reserve_slot(&mut chunks[current]);
    let chunk = &mut chunks[current];

    chunk.emit_op_u16(Op::LOCAL_SET, lines_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, lines_slot, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    host::emit(chunk, "ecma:array", "join", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    fs_path::emit_write_file(chunk, line);
    emit_write_or_throw(chunks, current, path_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_path_get_file_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("\\")), line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    host::emit(chunk, "ecma:string", "replaceAll", 3, line);
    paths::emit_file_name(chunk, line);
}

fn emit_normalized_path(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::String(Arc::from("\\")), line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    host::emit(chunk, "ecma:string", "replaceAll", 3, line);
}

pub fn emit_path_get_directory_name(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    paths::emit_directory(&mut chunks[current], line);
}

pub fn emit_path_get_file_name_without_extension(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    paths::emit_file_stem(&mut chunks[current], line);
}

pub fn emit_path_combine(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    paths::emit_combine(chunk, argc, line);
}

pub fn emit_path_change_extension(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let ext_slot = reserve_slot(chunk);
    let path_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, ext_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ext_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ext_slot, line);
    chunk.emit_end(line);

    paths::emit_change_extension(chunk, line);
}

pub fn emit_path_get_full_path(chunks: &mut [Chunk], current: usize, line: u32) {
    paths::emit_full_path(&mut chunks[current], line);
}

/// `Path.GetPathRoot(p)` — `"/"` for a rooted path, `""` otherwise.
///
/// `node:path.parse().root` answered this, which meant a `Path` call reaching
/// into Node's module graph to read one character. On the POSIX paths this
/// target has, the root IS the leading separator, so `IsPathRooted` already
/// decides it. Windows drive roots (`C:\`) are not modelled here — neither was
/// they by `parse` on this host.
pub fn emit_path_get_path_root(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    paths::emit_is_rooted(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    chunk.emit_else(line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_end(line);
}

pub fn emit_path_get_temp_file_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);

    paths::emit_temp_path(chunk, line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    push_const(chunk, Value::String(Arc::from("vybe-dotnet-")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    push_const(chunk, Value::F64(1000000.0), line);
    host::emit(chunk, "ecma:math", "random", 0, line);
    chunk.emit_op(Op::F64_MUL, line);
    host::emit(chunk, "ecma:math", "floor", 1, line);
    host::emit(chunk, "ecma:string", "toString", 1, line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    push_const(chunk, Value::String(Arc::from(".tmp")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    // `Path.GetTempFileName` CREATES the file — the name is only unique because
    // creating it reserves it. The success is dropped rather than thrown on:
    // the temp directory exists by construction here, so a failure means
    // something the caller cannot act on either way.
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("")), line);
    fs_path::emit_write_file(chunk, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
}

pub fn emit_path_get_random_file_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(1000000000000.0), line);
    host::emit(chunk, "ecma:math", "random", 0, line);
    chunk.emit_op(Op::F64_MUL, line);
    host::emit(chunk, "ecma:math", "floor", 1, line);
    host::emit(chunk, "ecma:string", "toString", 1, line);
    push_const(chunk, Value::String(Arc::from(".tmp")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
}

pub fn emit_path_get_invalid_file_name_chars(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    for ch in ["\"", "<", ">", "|", "\0", ":", "*", "?", "\\", "/"] {
        push_const(chunk, Value::String(Arc::from(ch)), line);
    }
    collections::emit_array_new(chunks, current, 10, line);
}

pub fn emit_path_get_invalid_path_chars(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    for ch in ["\"", "<", ">", "|", "\0"] {
        push_const(chunk, Value::String(Arc::from(ch)), line);
    }
    collections::emit_array_new(chunks, current, 5, line);
}

pub fn emit_path_has_extension(chunks: &mut [Chunk], current: usize, line: u32) {
    paths::emit_has_extension(&mut chunks[current], line);
}

pub fn emit_path_is_path_rooted(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    paths::emit_is_rooted(&mut chunks[current], line);
}

/// `Path.GetRelativePath(from, to)`.
///
/// ⚠THE ONE `node:path` CALL LEFT IN THIS FILE, and it is named rather than
/// quietly kept: relativising two paths is a real algorithm (split both, drop
/// the common prefix, emit one `..` per remaining `from` segment), not a
/// lookup, and `primitives::paths` has no equivalent to move it to. Writing a
/// second one here — in a .NET adapter, where every other language would then
/// need its own — is the per-language-restatement mistake.
///
/// It is not an INVENTED name and it is not filesystem access: `node:path` is a
/// real module and this is pure string work on paths that need not exist. The
/// fix is a `paths::emit_relative` for every language to share.
pub fn emit_path_get_relative_path(chunks: &mut [Chunk], current: usize, line: u32) {
    let relative = chunks[current].add_import("node:path", "relative");
    let chunk = &mut chunks[current];
    chunk.emit_call(relative, 2, line);
}

pub fn emit_path_trim_ending_directory_separator(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    host::emit(chunk, "ecma:string", "endsWith", 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("\\")), line);
    host::emit(chunk, "ecma:string", "endsWith", 2, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    host::emit(chunk, "ecma:string", "length", 1, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    host::emit(chunk, "ecma:string", "slice", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_directory_entries(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
    want_directories: bool,
) {
    let chunk = &mut chunks[current];
    let root_slot = reserve_slot(chunk);
    let entries_slot = reserve_slot(chunk);
    let idx_slot = reserve_slot(chunk);
    let entry_slot = reserve_slot(chunk);
    let full_path_slot = reserve_slot(chunk);
    let pattern_slot = reserve_slot(chunk);
    let allowed_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    if argc > 1 {
        chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    } else {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, root_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, root_slot, line);
    vybe_compiler::primitives::fs_path::emit_list_dir(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entries_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, entries_slot, idx_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, entry_slot, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_op_u16(Op::LOCAL_SET, allowed_slot, line);

    // `Path.Combine`, not Node's `resolve`: the entry name is always relative
    // to the directory just enumerated, and `resolve` would additionally
    // consult the process cwd for a root that is itself relative.
    chunk.emit_op_u16(Op::LOCAL_GET, root_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
    paths::emit_combine(chunk, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, full_path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, full_path_slot, line);
    fs_path::emit_is_dir(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    if !want_directories {
        vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    }

    let skip_push = chunk.emit_block(line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    if !want_directories {
        chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
        chunk.emit_else(line);

        chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
        push_const(chunk, Value::String(Arc::from("*")), line);
        host::emit(chunk, "ecma:string", "startsWith", 2, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        host::emit(chunk, "ecma:string", "slice", 2, line);
        host::emit(chunk, "ecma:string", "endsWith", 2, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_SET, allowed_slot, line);
        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, allowed_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
        chunk.emit_br_if(0, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, full_path_slot, line);
    collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.patch_block(skip_push);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_directory_get_files(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_directory_entries(chunks, current, argc, line, false);
}

pub fn emit_directory_get_directories(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_directory_entries(chunks, current, 1, line, true);
}

pub fn emit_directory_delete(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];

    if argc > 1 {
        chunk.emit_op(Op::DROP, line);
    }

    fs_path::emit_remove(chunk, line);
}
