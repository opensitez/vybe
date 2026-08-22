//! Python `open()` and the file object it returns.
//!
//! The profile used to map `open` → `wasi:filesystem:open` and
//! `close`/`readline`/`readlines` → `wasi:filesystem:close`/`readLine`/
//! `readLines` — **none of which the wasi shim registers** (it has `openFile`,
//! `closeFile`, `lineInput`). So every one of those was an
//! `Unresolved import` and Python had no working file I/O at all.
//!
//! Rather than chase per-handle host fns, `open()` returns a small object
//! carrying `__fpath` / `__fmode` / `__fdata` / `__fpos`, and the file methods
//! are adapters over the shim's whole-file primitives (`readFile`, `writeFile`,
//! `appendFile`) — wasi only, no node needed. Writes flush on every call, so a
//! missing `close()` never loses data and `with` needs no special support.
//!
//! `read`/`write`/`close` are shared value-method names, so each adapter probes
//! `__fpath` first and falls through to its previous behaviour for any other
//! receiver (a StringIO, a socket, …) rather than hijacking the name.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::{fs_path, ops, strings};

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

fn set_field(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let k = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(key)));
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

fn field_of(chunks: &mut [Chunk], current: usize, slot: u16, key: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let k = chunks[current].add_constant(vybe_runtime::Value::String(std::sync::Arc::from(key)));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

/// True when `slot` is one of our file objects (has `__fpath`).
fn emit_is_file_obj(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    field_of(chunks, current, slot, "__fpath", line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
}

/// Push `mode.includes(<c>)`.
fn emit_mode_has(chunks: &mut [Chunk], current: usize, mode: u16, c: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, mode, line);
    chunks[current].emit_string_const(c, line);
    call_import(chunks, current, "ecma:string", "includes", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// `open(path[, mode])` → file object. Read modes preload the contents; write
/// mode truncates immediately so an empty `with open(p,"w")` still clears the
/// file, matching CPython.
pub fn emit_open(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let path = base;
    let mode = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    } else {
        chunks[current].emit_string_const("r", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, mode, line);

    // "w" truncates up front.
    emit_mode_has(chunks, current, mode, "w", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    chunks[current].emit_string_const("", line);
    fs_path::emit_write_file(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    // Read modes preload; write-only files start with empty contents.
    let data = chunks[current].alloc_scratch(1);
    emit_mode_has(chunks, current, mode, "r", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_read_file(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);

    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    set_field(chunks, current, "__fpath", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mode, line);
    set_field(chunks, current, "__fmode", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    set_field(chunks, current, "__fdata", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_f64_const(0.0, line);
    set_field(chunks, current, "__fpos", line);
}

/// `f.write(s)` — append to the file and to the in-memory copy. Flushes every
/// call, so no `close()` is required for the bytes to land.
pub fn emit_write(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let f = base;
    let s = base + 1;

    emit_is_file_obj(chunks, current, f, line);
    chunks[current].emit_if_value(line);
    field_of(chunks, current, f, "__fpath", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    fs_path::emit_append_file(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);
    // Keep __fdata in step so a read-after-write on the same handle is correct.
    chunks[current].emit_op_u16(Op::LOCAL_GET, f, line);
    field_of(chunks, current, f, "__fdata", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set_field(chunks, current, "__fdata", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    chunks[current].emit_else(line);
    // Not one of ours — the previous whole-file host write.
    chunks[current].emit_op_u16(Op::LOCAL_GET, f, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    fs_path::emit_write_file(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// `f.read()` — everything from the current position on.
pub fn emit_read(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let f = base;

    emit_is_file_obj(chunks, current, f, line);
    chunks[current].emit_if_value(line);
    let data = chunks[current].alloc_scratch(1);
    field_of(chunks, current, f, "__fdata", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    field_of(chunks, current, f, "__fpos", line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    strings::emit_length(&mut chunks[current], line);
    strings::emit_substring(&mut chunks[current], line);
    // Position moves to EOF, so a second read() answers "" as CPython does.
    chunks[current].emit_op_u16(Op::LOCAL_GET, f, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    set_field(chunks, current, "__fpos", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, f, line);
    fs_path::emit_read_file(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// `f.readlines()` — line list, terminators kept, as CPython does.
pub fn emit_readlines(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let f = base;
    emit_is_file_obj(chunks, current, f, line);
    chunks[current].emit_if_value(line);
    field_of(chunks, current, f, "__fdata", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, f, line);
    fs_path::emit_read_file(&mut chunks[current], line);
    chunks[current].emit_end(line);
    // keepends=True — CPython's readlines keeps the terminators.
    chunks[current].emit_f64_const(1.0, line);
    crate::emitter::string_adapter::emit_splitlines(chunks, current, 2, line);
}

/// `f.close()` — writes already flushed, so this only marks the handle closed.
pub fn emit_close(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let f = base;
    emit_is_file_obj(chunks, current, f, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, f, line);
    chunks[current].emit_f64_const(1.0, line);
    set_field(chunks, current, "__fclosed", line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `f.readline()` — the next line INCLUDING its terminator, `""` at EOF.
pub fn emit_readline(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let f = base;
    let data = chunks[current].alloc_scratch(1);
    let pos = chunks[current].alloc_scratch(1);
    let nl = chunks[current].alloc_scratch(1);
    let end = chunks[current].alloc_scratch(1);

    field_of(chunks, current, f, "__fdata", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, data, line);
    field_of(chunks, current, f, "__fpos", line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pos, line);

    // nl = data.indexOf("\n", pos)
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_string_const("\n", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pos, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    call_import(chunks, current, "ecma:string", "indexOf", 3, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nl, line);

    // end = nl < 0 ? len(data) : nl + 1   (terminator is kept)
    chunks[current].emit_op_u16(Op::LOCAL_GET, nl, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nl, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, f, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    set_field(chunks, current, "__fpos", line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, data, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pos, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end, line);
    strings::emit_substring(&mut chunks[current], line);
}

/// `f.writelines(seq)` — concatenate and write; CPython adds no separators.
pub fn emit_writelines(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let f = base;
    let seq = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_GET, f, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, seq, line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "ecma:array", "join", 2, line);
    emit_write_from_stack(chunks, current, line);
}

/// Shared tail of `write`, with `[file, text]` already on the stack.
fn emit_write_from_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    let f = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, f, line);
    field_of(chunks, current, f, "__fpath", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    fs_path::emit_append_file(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, f, line);
    field_of(chunks, current, f, "__fdata", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set_field(chunks, current, "__fdata", line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `f.seek(n)` — byte offset within the buffered contents.
pub fn emit_seek(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    set_field(chunks, current, "__fpos", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
}

/// `f.tell()`.
pub fn emit_tell(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    field_of(chunks, current, base, "__fpos", line);
}

/// Build a fresh temp path from the walker-normalized `(prefix, suffix, dir)`
/// triple: `<dir or tmpdir>/<prefix>tmp<uuid><suffix>`. Stack: `[]` → `[path]`.
fn emit_temp_path(chunks: &mut [Chunk], current: usize, base: u16, argc: u8, line: u32) {
    let has = |i: u8| argc > i;
    // dir — empty string means "use the system temp dir".
    if has(2) {
        let d = base + 2;
        chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
        strings::emit_length(&mut chunks[current], line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, d, line);
        chunks[current].emit_else(line);
        call_import(chunks, current, "node:os", "tmpdir", 0, line);
        chunks[current].emit_end(line);
    } else {
        call_import(chunks, current, "node:os", "tmpdir", 0, line);
    }
    chunks[current].emit_string_const("/", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    if has(0) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
        ops::emit_dyn_add(&mut chunks[current], line);
    }
    chunks[current].emit_string_const("tmp", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    // A unique token for the temp NAME — not a canonical UUID, and nothing
    // downstream parses it as one. `wasi:random/random.uuid` was an invented
    // verb: `wasi:random@0.3.1` declares `get-random-bytes` and
    // `get-random-u64`, and nothing else. 64 bits of cryptographically strong
    // entropy is more than CPython's `tempfile`, which uses eight characters
    // from a 62-symbol alphabet (~47 bits).
    call_import(chunks, current, "wasi:random/random", "get-random-u64", 0, line);
    strings::emit_to_string(&mut chunks[current], line);
    ops::emit_dyn_add(&mut chunks[current], line);
    if has(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
        ops::emit_dyn_add(&mut chunks[current], line);
    }
}

/// `tempfile.gettempdir()`.
pub fn emit_gettempdir(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "node:os", "tmpdir", 0, line);
}

/// `tempfile.mkdtemp()` → a freshly created directory's path.
pub fn emit_mkdtemp(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let p = chunks[current].alloc_scratch(1);
    emit_temp_path(chunks, current, base, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, p, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    fs_path::emit_mkdir(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
}

/// `tempfile.NamedTemporaryFile(...)` — a file object that also carries `.name`,
/// which is the attribute callers read to get the path.
pub fn emit_named_temp_file(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // The walker flattened the keywords to (prefix, suffix, dir).
    let base = stash_args(chunks, current, argc, line);
    let p = chunks[current].alloc_scratch(1);
    emit_temp_path(chunks, current, base, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, p, line);
    // Create it empty so `os.path.exists` is true before anything is written.
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    chunks[current].emit_string_const("", line);
    fs_path::emit_write_file(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    set_field(chunks, current, "__fpath", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    set_field(chunks, current, "name", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("w+", line);
    set_field(chunks, current, "__fmode", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    set_field(chunks, current, "__fdata", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_f64_const(0.0, line);
    set_field(chunks, current, "__fpos", line);
}

/// `tempfile.gettempprefix()`.
pub fn emit_tmp_prefix(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_string_const("tmp", line);
}

/// `tempfile.mkstemp()` → `(fd, path)`. There are no real file descriptors
/// here, so the fd slot is a stable placeholder; callers use the path.
pub fn emit_mkstemp(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let p = chunks[current].alloc_scratch(1);
    emit_temp_path(chunks, current, base, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, p, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    chunks[current].emit_string_const("", line);
    fs_path::emit_write_file(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 2, line);
}

/// `os.path.samefile(a, b)` — the shim has no inode, so compare resolved paths.
pub fn emit_samefile(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    vybe_compiler::primitives::paths::emit_full_path(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    vybe_compiler::primitives::paths::emit_full_path(&mut chunks[current], line);
    call_import(chunks, current, "wasm:js-string", "equals", 2, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}
