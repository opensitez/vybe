//! Python `glob`, `linecache` and `inspect`.
//!
//! All three are file/object introspection over machinery that already
//! exists: the shared glob matcher (`primitives/strings.rs`, which php's
//! `fnmatch` also binds), the shared filesystem lowering
//! (`primitives/fs_path.rs`) and the shared object view
//! (`primitives/reflection.rs`).

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::{
    collections, fs_path, ops, reflection, strings, tuples,
};

use super::adapter_util::{call_import, lget, lset, stash_exact};

/// `glob.glob(pattern)` — the entries of the pattern's directory that match
/// its final component, each with the directory prefix put back.
///
/// A leading dot is only matched by a pattern that also starts with one,
/// which is the one rule `fnmatch` does not carry (CPython's `glob` applies
/// it on top of the same matcher).
pub fn emit_glob(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let pattern = chunks[current].alloc_scratch(6);
    let dir = pattern + 1;
    let prefix = pattern + 2;
    let leaf = pattern + 3;
    let entries = pattern + 4;
    let out = pattern + 5;

    lget(&mut chunks[current], base, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    lset(&mut chunks[current], pattern, line);

    // Split at the last separator: everything before it is the directory to
    // list, everything after is what each entry has to match.
    let cut = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], pattern, line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "ecma:string", "lastIndexOf", 2, line);
    lset(&mut chunks[current], cut, line);

    lget(&mut chunks[current], cut, line);
    call_import(chunks, current, "wasm:js-number", "toI32", 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const(".", line);
    lset(&mut chunks[current], dir, line);
    chunks[current].emit_string_const("", line);
    lset(&mut chunks[current], prefix, line);
    lget(&mut chunks[current], pattern, line);
    lset(&mut chunks[current], leaf, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], pattern, line);
    chunks[current].emit_i32_const(0, line);
    lget(&mut chunks[current], cut, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    lset(&mut chunks[current], dir, line);
    lget(&mut chunks[current], pattern, line);
    chunks[current].emit_i32_const(0, line);
    lget(&mut chunks[current], cut, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    lset(&mut chunks[current], prefix, line);
    lget(&mut chunks[current], pattern, line);
    lget(&mut chunks[current], cut, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lget(&mut chunks[current], pattern, line);
    strings::emit_length(&mut chunks[current], line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    lset(&mut chunks[current], leaf, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], dir, line);
    fs_path::emit_list_dir(&mut chunks[current], line);
    lset(&mut chunks[current], entries, line);
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out, line);

    let i = chunks[current].alloc_scratch(2);
    let name = i + 1;
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);
    let loop_id = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], entries, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], entries, line);
    lget(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], name, line);

    // matches = glob_match(name, leaf) && (leaf starts "." || name does not)
    lget(&mut chunks[current], name, line);
    lget(&mut chunks[current], leaf, line);
    // Case-SENSITIVE: `os.path.normcase` is identity on POSIX, so python's
    // glob and fnmatch are the same match (the profile's `__glob_match` says
    // the same thing).
    strings::emit_glob_match(
        chunks,
        current,
        vybe_compiler::primitives::strings::GlobOptions::exact(),
        line,
    );
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    lget(&mut chunks[current], leaf, line);
    chunks[current].emit_string_const(".", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    lget(&mut chunks[current], name, line);
    chunks[current].emit_string_const(".", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], prefix, line);
    lget(&mut chunks[current], name, line);
    let concat = chunks[current].add_import("wasm:js-string", "concat");
    chunks[current].emit_call(concat, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_id, line);

    lget(&mut chunks[current], out, line);
}

/// `glob.escape(path)` — wrap the three special characters in `[]`.
pub fn emit_glob_escape(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    lget(&mut chunks[current], base, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    for (from, to) in [("*", "[*]"), ("?", "[?]")] {
        chunks[current].emit_string_const(from, line);
        chunks[current].emit_string_const(to, line);
        call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    }
}

/// `linecache.getline(path, lineno)` — the 1-based line WITH its terminator,
/// or `""` when the line does not exist (CPython never raises here).
pub fn emit_getline(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let lines = chunks[current].alloc_scratch(2);
    let index = lines + 1;

    lget(&mut chunks[current], base, line);
    fs_path::emit_read_file(&mut chunks[current], line);
    chunks[current].emit_f64_const(1.0, line);
    crate::emitter::string_adapter::emit_splitlines(chunks, current, 2, line);
    lset(&mut chunks[current], lines, line);

    lget(&mut chunks[current], base + 1, line);
    call_import(chunks, current, "wasm:js-number", "toI32", 1, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(&mut chunks[current], index, line);

    lget(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    lget(&mut chunks[current], index, line);
    lget(&mut chunks[current], lines, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], lines, line);
    lget(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
}

/// `linecache.getlines(path)` — every line, terminators kept.
pub fn emit_getlines(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    lget(&mut chunks[current], base, line);
    fs_path::emit_read_file(&mut chunks[current], line);
    chunks[current].emit_f64_const(1.0, line);
    crate::emitter::string_adapter::emit_splitlines(chunks, current, 2, line);
}

/// `linecache.clearcache()` / `checkcache()` — nothing is cached, so there is
/// nothing to drop; both return None in CPython too.
pub fn emit_linecache_none(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// A class object carries a `prototype` (`primitives/classes.rs` sets it on
/// every class it emits); a plain function does not. That is the one runtime
/// difference between the two, both being `typeof "function"`.
/// Stack: `[]` → `[i32]`, reading the value in `slot`.
fn emit_has_prototype(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    chunks[current].emit_string_const("prototype", line);
    let get = chunks[current].add_import("ecma:object", "get");
    chunks[current].emit_call(get, 2, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
}

/// Stack: `[]` → `[i32]` — is the value in `slot` callable at all?
fn emit_is_function_value(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    reflection::emit_typeof(chunks, current, line);
    chunks[current].emit_string_const("function", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// `inspect.isclass(x)` — callable AND carrying a prototype.
pub fn emit_isclass(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    emit_is_function_value(chunks, current, base, line);
    emit_has_prototype(chunks, current, base, line);
    chunks[current].emit_op(Op::I32_AND, line);
    let from_i32 = chunks[current].add_import("wasm:js-boolean", "fromI32");
    chunks[current].emit_call(from_i32, 1, line);
}

/// `inspect.isfunction(x)` — callable and NOT a class.
pub fn emit_isfunction(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    emit_is_function_value(chunks, current, base, line);
    emit_has_prototype(chunks, current, base, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    let from_i32 = chunks[current].add_import("wasm:js-boolean", "fromI32");
    chunks[current].emit_call(from_i32, 1, line);
}

/// `inspect.isroutine(x)` / `isbuiltin(x)` — callable.
pub fn emit_iscallable(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    emit_is_function_value(chunks, current, base, line);
    let from_i32 = chunks[current].add_import("wasm:js-boolean", "fromI32");
    chunks[current].emit_call(from_i32, 1, line);
}

/// `inspect.getmembers(obj)` → `[(name, value)]`, sorted by name as CPython
/// sorts it. The pairs come from the shared object view, so a class and an
/// instance are handled by the same lowering.
pub fn emit_getmembers(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    let pairs = chunks[current].alloc_scratch(3);
    let i = pairs + 1;
    let out = pairs + 2;

    lget(&mut chunks[current], base, line);
    reflection::emit_object_view(chunks, current, reflection::ObjectKeysMode::Entries, line);
    lset(&mut chunks[current], pairs, line);
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out, line);

    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);
    let loop_id = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], pairs, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    // Each entry is a 2-element array; Python wants a real tuple.
    let entry = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], pairs, line);
    lget(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], entry, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], entry, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    lget(&mut chunks[current], entry, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    tuples::emit_tuple(chunks, current, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_id, line);

    lget(&mut chunks[current], out, line);
}
