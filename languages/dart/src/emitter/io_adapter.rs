//! `dart:io` filesystem — Rust inline opcode emitters.
//!
//! A `File`/`Directory`/`Link` is lowered by the walker to a record
//! `{ path, __dart_io: "file" }`, so every method here starts by reading
//! `path` off the receiver and then calls an EXISTING `wasi:filesystem`
//! host function. No host fn is added: `fs.rs` already registers
//! `readFile`, `readFileBytes`, `writeFile`, `appendFile`, `exists`,
//! `isFile`, `isDir`, `remove`, `listDir`, `mkdir`, `fileSize`, `rename`
//! and `copy`, which is the whole synchronous surface these methods need.
//!
//! Value-method argc INCLUDES the receiver, and arguments sit ABOVE it on
//! the stack — so an N-argument method pops N values, then the receiver.

use std::sync::Arc;
use vybe_compiler::primitives::{errors, ops, strings};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

/// Must match `walker::DART_IO_KIND_KEY`.
pub const DART_IO_KIND_KEY: &str = "__dart_io";

fn slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn string_key(chunk: &mut Chunk, key: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(key)))
}

/// Pop `argc - 1` arguments into slots (last argument first, since it is on
/// top), then pop the receiver and leave its `path` in a slot.
///
/// Returns `(path_slot, arg_slots)` with `arg_slots` in SOURCE order.
fn take_receiver_path(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> (u16, Vec<u16>) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();

    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);

    let path_slot = slot(&mut chunks[current]);
    let path_key = string_key(&mut chunks[current], "path");
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, path_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path_slot, line);

    (path_slot, arg_slots)
}

fn call_fs(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("wasi:filesystem", name);
    chunks[current].emit_call(idx, argc, line);
}

/// `wasi:filesystem.exists(path)` — leaves a bool on the stack.
fn push_exists(chunks: &mut [Chunk], current: usize, path_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "exists", 1, line);
}

/// Dart throws `FileSystemException` for a missing path; the host `readFile`
/// instead returns the string `"Error: …"`, which would silently become the
/// file's contents. Guard first so the failure is a real throw.
///
/// `require` names the host predicate: a File read demands `isFile`, not mere
/// existence — `File(someDirectory).readAsStringSync()` throws in Dart, and
/// guarding on `exists` let a directory through and returned the host's error
/// text as if it were file contents.
fn throw_if_not(
    chunks: &mut [Chunk],
    current: usize,
    path_slot: u16,
    require: &str,
    message: &str,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, require, 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, message, line);
    chunks[current].emit_end(line);
}

/// `throw FileSystemException("<message><path>")`.
///
/// Routed through the crate's one Dart-exception builder rather than calling
/// `emit_exception_new_finalize` directly: that helper's contract is
/// `[obj, obj, msg]` on the stack, and handing it only the message produced an
/// object `on FileSystemException` would not catch — the throw escaped every
/// `try`.
fn emit_filesystem_throw(
    chunks: &mut [Chunk],
    current: usize,
    path_slot: u16,
    message: &str,
    line: u32,
) {
    chunks[current].emit_string_const(message, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    strings::emit_concat(&mut chunks[current], 2, line);
    crate::emitter::string_adapter::emit_dart_exception_new(
        chunks,
        current,
        1,
        "FileSystemException",
        &["FileSystemException", "IOException", "Exception"],
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
}

/// `file.readAsStringSync()`
pub fn emit_read_as_string_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(chunks, current, path_slot, "isFile", "Cannot open file, path = ", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "readFile", 1, line);
}

/// `file.readAsBytesSync()`
pub fn emit_read_as_bytes_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(chunks, current, path_slot, "isFile", "Cannot open file, path = ", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "readFileBytes", 1, line);
}

/// `file.readAsLinesSync()` — Dart treats a single trailing newline as a
/// terminator rather than as a final empty line, so ONE trailing `\n` is
/// removed before splitting. Trimming instead would also eat deliberate blank
/// lines at the end of the file.
pub fn emit_read_as_lines_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(chunks, current, path_slot, "isFile", "Cannot open file, path = ", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "readFile", 1, line);

    let text_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);

    let ends_with = chunks[current].add_import("ecma:string", "endsWith");
    let str_slice = chunks[current].add_import("ecma:string", "slice");
    let str_len = chunks[current].add_import("ecma:string", "length");

    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_string_const("\n", line);
    chunks[current].emit_call(ends_with, 2, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_call(str_len, 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_call(str_slice, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_string_const("\n", line);
    strings::emit_split(&mut chunks[current], line);
}

/// A write whose host call reports failure must THROW: Dart raises
/// `FileSystemException` when the path is a directory or unwritable, and the
/// host functions report that as `false`. Leaving the bool on the stack meant
/// `File('/').writeAsStringSync(…)` silently succeeded-as-false.
fn emit_write_via(chunks: &mut [Chunk], current: usize, argc: u8, host_fn: &str, line: u32) {
    let (path_slot, args) = take_receiver_path(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    match args.first() {
        Some(data) => chunks[current].emit_op_u16(Op::LOCAL_GET, *data, line),
        None => chunks[current].emit_string_const("", line),
    }
    call_fs(chunks, current, host_fn, 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Cannot open file, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `file.writeAsStringSync(contents)` — truncating write.
pub fn emit_write_as_string_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_write_via(chunks, current, argc, "writeFile", line);
}

/// `file.writeAsStringSync(contents, mode: FileMode.append)`
pub fn emit_append_as_string_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_write_via(chunks, current, argc, "appendFile", line);
}

/// `handle.existsSync()` — a File must be a FILE and a Directory a DIRECTORY,
/// which is why the record carries its kind: `File('some_dir').existsSync()`
/// is false in Dart even though the path exists.
pub fn emit_exists_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);

    let path_key = string_key(&mut chunks[current], "path");
    let kind_key = string_key(&mut chunks[current], DART_IO_KIND_KEY);
    let path_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, path_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, kind_key, line);
    chunks[current].emit_string_const("directory", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "isDir", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "isFile", 1, line);
    chunks[current].emit_end(line);
}

/// `handle.deleteSync()` — `remove` covers both a file and a directory.
pub fn emit_delete_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(chunks, current, path_slot, "exists", "Deletion failed, path = ", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "remove", 1, line);
}

/// `file.lengthSync()`
pub fn emit_length_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(chunks, current, path_slot, "isFile", "Cannot retrieve length, path = ", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "fileSize", 1, line);
}

/// `handle.createSync()` — a Directory makes the directory; a File makes an
/// empty file, which `writeFile` with empty contents already does. Creating a
/// file that exists must NOT truncate it, so the write is guarded.
pub fn emit_create_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);

    let path_key = string_key(&mut chunks[current], "path");
    let kind_key = string_key(&mut chunks[current], DART_IO_KIND_KEY);
    let path_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, path_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, kind_key, line);
    chunks[current].emit_string_const("directory", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "mkdir", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    push_exists(chunks, current, path_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunks[current].emit_string_const("", line);
    call_fs(chunks, current, "writeFile", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
}

/// `handle.renameSync(newPath)` / `file.copySync(newPath)` — both return a
/// NEW handle of the same kind at the destination, which is why the receiver's
/// kind is copied onto the result rather than assumed to be a file.
fn emit_relocate(chunks: &mut [Chunk], current: usize, argc: u8, host_fn: &str, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);

    let path_key = string_key(&mut chunks[current], "path");
    let kind_key = string_key(&mut chunks[current], DART_IO_KIND_KEY);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, path_key, line);
    match arg_slots.first() {
        Some(dest) => chunks[current].emit_op_u16(Op::LOCAL_GET, *dest, line),
        None => chunks[current].emit_string_const("", line),
    }
    call_fs(chunks, current, host_fn, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    // The result handle: same kind, destination path.
    let out_slot = slot(&mut chunks[current]);
    chunks[current].emit_op(Op::STRUCT_NEW_DEFAULT, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    match arg_slots.first() {
        Some(dest) => chunks[current].emit_op_u16(Op::LOCAL_GET, *dest, line),
        None => chunks[current].emit_string_const("", line),
    }
    chunks[current].emit_op_u16(Op::STRUCT_SET, path_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, kind_key, line);
    chunks[current].emit_op_u16(Op::STRUCT_SET, kind_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_rename_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_relocate(chunks, current, argc, "rename", line);
}

pub fn emit_copy_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_relocate(chunks, current, argc, "copy", line);
}

/// `directory.listSync()` — Dart yields File/Directory HANDLES, not strings,
/// so each host path is wrapped back into a record whose kind is decided by
/// asking the filesystem what the entry actually is.
pub fn emit_list_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(chunks, current, path_slot, "isDir", "Directory listing failed, path = ", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "listDir", 1, line);
}
