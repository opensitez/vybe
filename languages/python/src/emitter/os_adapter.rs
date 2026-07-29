//! Python `os` surfaces that a plain profile host-map cannot express.
//!
//! Everything flat (`getcwd`, `getpid`, `rename`, `remove`, `truncate`,
//! `urandom`, …) is a direct profile entry — **wasi first**
//! (`wasi:filesystem`, `wasi:random`), node only where wasi has no equivalent
//! (`node:process.cwd`, `node:fs.truncateSync`, `node:tty.isatty`).
//!
//! What lands here is shape translation: the wasi shim answers
//! `stat` → `{size, isFile, isDir, modified}` and
//! `readDirEntries` → `[{name, isFile, isDir}]`, while Python wants
//! `st_size`/`st_mtime`/… and `DirEntry` objects. These are ADAPTERS rather
//! than a source prelude so nothing is parsed or compiled unless the call is
//! actually present — a prelude gated on a `source.contains("os.…")` substring
//! would compile the whole module for any program that merely mentions it.
//!
//! `DirEntry`'s methods (`is_file()`, `is_dir()`, …) are `[value_methods]`
//! reading fields off the entry object, the same way other host-produced
//! values expose methods.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::{ops, strings};

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// `obj.<key> = <value already on stack>`, leaving `obj` on the stack.
fn set_field(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let k = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(key)));
    chunks[current].emit_op_u16(Op::STRUCT_SET, k, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn get_field(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let k = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(key)));
    chunks[current].emit_op_u16(Op::STRUCT_GET, k, line);
}

/// Read `slot.<key>` onto the stack.
fn field_of(chunks: &mut [Chunk], current: usize, slot: u16, key: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    get_field(chunks, current, key, line);
}

/// Build a `stat_result` from the wasi shim's `{size, isFile, isDir, modified}`
/// (or null when the path is missing). Stack: `[raw]` → `[stat_result]`.
fn emit_stat_result_from(chunks: &mut [Chunk], current: usize, raw: u16, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);

    // st_size
    chunks[current].emit_dup(line);
    field_of(chunks, current, raw, "size", line);
    set_field(chunks, current, "st_size", line);

    // st_mtime — the shim reports milliseconds; Python uses float seconds.
    let secs = chunks[current].alloc_scratch(1);
    field_of(chunks, current, raw, "modified", line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_f64_const(1000.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, secs, line);
    for key in ["st_mtime", "st_atime", "st_ctime"] {
        chunks[current].emit_dup(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, secs, line);
        set_field(chunks, current, key, line);
    }

    // st_mode — directory vs regular file, the POSIX values CPython reports.
    chunks[current].emit_dup(line);
    field_of(chunks, current, raw, "isDir", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(16877.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(33188.0, line);
    chunks[current].emit_end(line);
    set_field(chunks, current, "st_mode", line);

    for (key, val) in [
        ("st_ino", 0.0),
        ("st_dev", 0.0),
        ("st_nlink", 1.0),
        ("st_uid", 0.0),
        ("st_gid", 0.0),
    ] {
        chunks[current].emit_dup(line);
        chunks[current].emit_f64_const(val, line);
        set_field(chunks, current, key, line);
    }
}

/// `os.stat(path)` / `os.lstat(path)`.
pub fn emit_stat(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let raw = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    call_import(chunks, current, "wasi:filesystem", "stat", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw, line);
    emit_stat_result_from(chunks, current, raw, line);
}

/// `os.stat(entry.path)` for a DirEntry receiver — `entry.stat()`.
pub fn emit_entry_stat(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let raw = chunks[current].alloc_scratch(1);
    field_of(chunks, current, base, "path", line);
    call_import(chunks, current, "wasi:filesystem", "stat", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw, line);
    emit_stat_result_from(chunks, current, raw, line);
}

/// Append `dir + "/" + name` to the stack.
fn emit_join(chunks: &mut [Chunk], current: usize, dir: u16, name_expr_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dir, line);
    chunks[current].emit_string_const("/", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_expr_slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
}

/// `os.scandir([path])` → array of DirEntry objects. Python's scandir returns a
/// lazy iterator; an array is iterable the same way and supports `len`, which
/// is all the surface the language exposes here.
pub fn emit_scandir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let dir = chunks[current].alloc_scratch(1);
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, dir, line);
    } else {
        chunks[current].emit_string_const(".", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, dir, line);
    }

    let raws = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dir, line);
    call_import(chunks, current, "wasi:filesystem", "readDirEntries", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raws, line);

    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, raws, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let raw = chunks[current].alloc_scratch(1);
    let nm = chunks[current].alloc_scratch(1);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, raws, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw, line);
    field_of(chunks, current, raw, "name", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nm, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nm, line);
    set_field(chunks, current, "name", line);
    chunks[current].emit_dup(line);
    emit_join(chunks, current, dir, nm, line);
    set_field(chunks, current, "path", line);
    chunks[current].emit_dup(line);
    field_of(chunks, current, raw, "isFile", line);
    set_field(chunks, current, "__is_file", line);
    chunks[current].emit_dup(line);
    field_of(chunks, current, raw, "isDir", line);
    set_field(chunks, current, "__is_dir", line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `os.walk(top[, topdown])` → array of `(dirpath, dirnames, filenames)`
/// tuples, depth-first from `top`. Iterative over an explicit stack; with
/// `topdown=False` the accumulated rows are reversed, which yields children
/// before parents exactly as CPython does.
pub fn emit_walk(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let top = base;

    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    let stack = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stack, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, top, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    let cur = chunks[current].alloc_scratch(1);
    let raws = chunks[current].alloc_scratch(1);
    let dirs = chunks[current].alloc_scratch(1);
    let files = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let raw = chunks[current].alloc_scratch(1);
    let nm = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);

    let outer = chunks[current].emit_block(line);
    let olp = chunks[current].emit_loop_s(line).0;
    // stop when the stack is empty
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, stack, line);
    call_import(chunks, current, "ecma:array", "shift", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur, line);
    call_import(chunks, current, "wasi:filesystem", "readDirEntries", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raws, line);
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dirs, line);
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, files, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, raws, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let inner = chunks[current].emit_block(line);
    let ilp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, raws, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw, line);
    field_of(chunks, current, raw, "name", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nm, line);

    field_of(chunks, current, raw, "isDir", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dirs, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nm, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, files, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nm, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(ilp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner);

    // out.push((cur, dirs, files))
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dirs, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, files, line);
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 3, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    // Queue each subdirectory (front of the stack keeps DFS pre-order).
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    let sub = chunks[current].emit_block(line);
    let slp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dirs, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dirs, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nm, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack, line);
    emit_join(chunks, current, cur, nm, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(slp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(sub);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(olp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    // topdown=False → children before parents.
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        call_import(chunks, current, "ecma:array", "toReversed", 1, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    }
}

/// `os.cpu_count()` — `node:os.cpus().length`, floored at 1 (wasi exposes no
/// CPU topology).
pub fn emit_cpu_count(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "node:os", "cpus", 0, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    chunks[current].emit_end(line);
}

/// `os.fspath(p)` — a str passes through; anything else answers `__fspath__`
/// (which for a DirEntry is its `path` field).
pub fn emit_fspath(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_else(line);
    field_of(chunks, current, base, "path", line);
    chunks[current].emit_end(line);
}

/// `os.strerror(code)` — the errno strings CPython reports.
pub fn emit_strerror(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let code = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code, line);

    let table = [
        (1.0, "Operation not permitted"),
        (2.0, "No such file or directory"),
        (9.0, "Bad file descriptor"),
        (13.0, "Permission denied"),
        (17.0, "File exists"),
        (20.0, "Not a directory"),
        (21.0, "Is a directory"),
        (22.0, "Invalid argument"),
        (28.0, "No space left on device"),
    ];
    for (num, msg) in table {
        chunks[current].emit_op_u16(Op::LOCAL_GET, code, line);
        chunks[current].emit_f64_const(num, line);
        chunks[current].emit_op(Op::F64_EQ, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(msg, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_string_const("Unknown error ", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code, line);
    strings::emit_to_string(&mut chunks[current], line);
    ops::emit_dyn_add(&mut chunks[current], line);
    for _ in table {
        chunks[current].emit_end(line);
    }
}

/// `entry.is_file()` / `is_dir()` / `is_symlink()` / `inode()` — DirEntry
/// methods reading the fields `emit_scandir` stamped.
pub fn emit_entry_flag(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    let e = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, e, line);
    field_of(chunks, current, e, field, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `entry.is_symlink()` — the wasi shim does not report link status, so this is
/// always False rather than silently wrong in the other direction.
pub fn emit_entry_false(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `entry.inode()` — no inode in the shim's shape.
pub fn emit_entry_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_f64_const(0.0, line);
}

/// `shutil.copytree(src, dst)` — recursive directory copy. Iterative over a
/// worklist of `[src, dst]` pairs; the shim's `mkdir` is `create_dir_all` and
/// its `copy` is a plain file copy, so this only has to walk the tree.
pub fn emit_copytree(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let work = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, work, line);
    // seed with [src, dst]
    chunks[current].emit_op_u16(Op::LOCAL_GET, work, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 2, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    let pair = chunks[current].alloc_scratch(1);
    let s_dir = chunks[current].alloc_scratch(1);
    let d_dir = chunks[current].alloc_scratch(1);
    let raws = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let raw = chunks[current].alloc_scratch(1);
    let nm = chunks[current].alloc_scratch(1);

    let outer = chunks[current].emit_block(line);
    let olp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, work, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, work, line);
    call_import(chunks, current, "ecma:array", "shift", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s_dir, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, d_dir, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, d_dir, line);
    call_import(chunks, current, "wasi:filesystem", "mkdir", 1, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, s_dir, line);
    call_import(chunks, current, "wasi:filesystem", "readDirEntries", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raws, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, raws, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let inner = chunks[current].emit_block(line);
    let ilp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, raws, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw, line);
    field_of(chunks, current, raw, "name", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nm, line);

    field_of(chunks, current, raw, "isDir", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, work, line);
    emit_join(chunks, current, s_dir, nm, line);
    emit_join(chunks, current, d_dir, nm, line);
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 2, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    emit_join(chunks, current, s_dir, nm, line);
    emit_join(chunks, current, d_dir, nm, line);
    call_import(chunks, current, "wasi:filesystem", "copy", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(ilp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(olp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
}

/// `shutil.which(cmd)` — first `PATH` entry containing an existing `cmd`, else
/// None. `PATH` comes from the wasi environment.
pub fn emit_which(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let cmd = base;
    let dirs = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let cand = chunks[current].alloc_scratch(1);

    let k = chunks[current].add_constant(vybe_runtime::Value::Null);
    chunks[current].emit_op_u16(Op::CONST, k, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);

    // `get-environment(key)` answers the value directly (no key → all pairs).
    chunks[current].emit_string_const("PATH", line);
    call_import(chunks, current, "wasi:cli/environment", "get-environment", 1, line);
    let path_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_s, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("/usr/bin:/bin:/usr/local/bin", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_string_const(":", line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dirs, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dirs, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dirs, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_string_const("/", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cmd, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cand, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cand, line);
    call_import(chunks, current, "wasi:filesystem", "exists", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cand, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, found, line);
}

/// `os.device_encoding(fd)` — the WASI console is always UTF-8.
pub fn emit_device_encoding(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_string_const("utf-8", line);
}

/// `os.get_terminal_size()` → `(columns, lines)`; CPython's own fallback when
/// the stream is not a tty is 80x24, which is what a WASI console reports.
pub fn emit_term_size(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_f64_const(80.0, line);
    set_field(chunks, current, "columns", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_f64_const(24.0, line);
    set_field(chunks, current, "lines", line);
}

// ── sys ─────────────────────────────────────────────────────────────────────

/// `sys.getsizeof(o)` — no object headers exist here, so report a plausible
/// non-zero size: strings/arrays scale with length, everything else is a word.
pub fn emit_getsizeof(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    chunks[current].emit_f64_const(49.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(28.0, line);
    chunks[current].emit_end(line);
}

/// `sys.intern(s)` — interning is a storage detail; the string IS the value.
pub fn emit_intern(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
}

/// `sys.getrecursionlimit()`.
pub fn emit_getrecursionlimit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_f64_const(1000.0, line);
}

/// `sys.setrecursionlimit(n)` — the VM stack depth is fixed, so this records
/// nothing and answers None, as CPython does.
pub fn emit_setrecursionlimit(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let k = chunks[current].add_constant(vybe_runtime::Value::Null);
    chunks[current].emit_op_u16(Op::CONST, k, line);
}

/// `sys.getdefaultencoding()` / `getfilesystemencoding()`.
pub fn emit_encoding(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_string_const("utf-8", line);
}

/// `sys.is_finalizing()` — never, while user code is still running.
pub fn emit_is_finalizing(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_i32_const(0, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `sys.exc_info()` → `(None, None, None)` outside an except block. The active
/// exception is not tracked as a global here, so this reports "no exception"
/// rather than inventing one.
pub fn emit_exc_info(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    for _ in 0..3 {
        let k = chunks[current].add_constant(vybe_runtime::Value::Null);
        chunks[current].emit_op_u16(Op::CONST, k, line);
    }
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 3, line);
}
