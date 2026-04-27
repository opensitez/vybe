//! WASM Standard Library — pure bytecode implementations of common functions.
//!
//! These are compiled to Chunk objects that get linked into every program.
//! On Vybe VM, the compiler can skip these and use host calls instead (faster).
//! On any other WASM runtime, these provide the same functionality portably.
//!
//! Functions:
//!   range(start, stop, step) → array
//!   sorted(array) → array (insertion sort)
//!   reversed(array) → array
//!   enumerate(array) → array of [i, val]
//!   zip(a, b) → array of [a_i, b_i]
//!   sum(array) → number
//!   min_val(array) → value
//!   max_val(array) → value
//!   to_str(value) → string (via convert import)
//!   string_is_null_or_empty(str) → bool
//!   string_is_null_or_whitespace(str) → bool
//!   str_insert(str, index, value) → string
//!   str_remove_start(str, start) → string
//!   str_remove_range(str, start, count) → string
//!   pow(base, exp) → number (repeated multiplication fallback)
//!   sin/cos/tan/asin/acos/atan/atan2/log/log10/exp/sign/clamp → number

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

// ── Generic polyglot polyfill helper ─────────────────────────────────
//
// `build_polyfill(source, language, export_name)` compiles a snippet
// of source code in any registered language and extracts a single
// named export as a stdlib Chunk. The result slots into `MAPPINGS` in
// `bundle.rs` exactly the same as a hand-built `build_pow`-style chunk.
//
// The mechanism: dispatch the source through the appropriate language
// parser, run the standard `Compiler::with_profile` over the AST, then
// pluck the chunk whose name matches `export_name` from the result.
// Parse / compile failures abort the build (panicking via `expect`)
// since polyfill sources ship with the binary — any breakage here
// is a vybex build bug, not a runtime error.
//
// This lets polyfills live in whichever source language is most
// natural — `sprintf.js`, `format.js`, `phpIncrement.php`,
// `vbFormat.vb`, etc. — instead of being hand-emitted as Rust opcode
// calls. Same final artifact: a Chunk that bundles into compiled
// programs and surfaces as a `__vybe_*` global.

/// Compile a polyfill source written in the given language, extracting
/// the chunk for the named export. The export name should match the
/// function declared in the source (e.g. `export function sprintf(...)`
/// in JS, `function sprintf(...) end function` in VB, etc.).
///
/// Errors include the language and export name in the message so build
/// failures point at the offending polyfill file.
pub(crate) fn build_polyfill(
    imports: &mut Chunk,
    source: &str,
    language: &str,
    export_name: &str,
) -> Chunk {
    let lang = crate::languages::find_by_name(language)
        .unwrap_or_else(|| panic!(
            "polyfill build: unknown language {:?} (registered: vb js pascal csharp \
             python php ruby dart cobol fortran)", language));
    let module = (lang.parse)(source)
        .unwrap_or_else(|e| panic!(
            "polyfill build: parse {:?}.{:?} failed: {}", language, export_name, e));
    let profile = crate::profile::parse_profile((lang.profile_source)())
        .unwrap_or_else(|e| panic!(
            "polyfill build: profile {:?} parse failed: {}", language, e));
    // Recursion guard so the inner compile pipeline skips its own
    // `finalize_with_stdlib` step — that would call back here and
    // recurse forever. Re-entrancy on the same thread is the only
    // failure mode and vybex build-time compilation is single-threaded.
    let polyfill_chunks = with_polyfill_guard(|| {
        crate::compiler::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap_or_else(|e| panic!(
                "polyfill build: compile {:?}.{:?} failed: {}", language, export_name, e))
    });

    // Merge the polyfill's module-level imports (which the JS compiler
    // wrote to its own chunks[0]) into the user program's imports
    // chunk, building a poly_idx → user_idx remap. Then walk the
    // function chunk's bytecode and rewrite every `CALL_IMPORT` operand
    // through the remap so runtime dispatch hits the right slot in the
    // user program's import table.
    let polyfill_script = polyfill_chunks.first()
        .unwrap_or_else(|| panic!("polyfill {}.{}: no chunks compiled", language, export_name));
    let remap: Vec<u16> = polyfill_script.imports.iter()
        .map(|imp| imports.add_import(imp.module.clone(), imp.name.clone()))
        .collect();

    let mut chunk = polyfill_chunks.into_iter()
        .find(|c| c.name == export_name)
        .unwrap_or_else(|| panic!(
            "polyfill build: export {:?} not found in {} source (chunks compiled, \
             but no chunk has that name — check the function is declared at \
             top level and exported)", export_name, language));

    if !remap.is_empty() {
        relocate_call_import_operands(&mut chunk, &remap);
    }
    chunk
}

/// Walk a chunk's bytecode and rewrite every `CALL_IMPORT` u16 operand
/// using `remap[poly_idx] = user_idx`. Uses the `OperandFormat` table to
/// safely advance past variable-length opcodes — same logic the
/// disassembler uses, so it stays aligned automatically.
fn relocate_call_import_operands(chunk: &mut Chunk, remap: &[u16]) {
    use vybe_bytecode::opcode::{Op, OperandFormat};
    let mut offset = 0;
    while offset + 1 < chunk.code.len() {
        let prefix = chunk.code[offset];
        let sub = chunk.code[offset + 1];
        let op = match Op::decode(prefix, sub) {
            Some(op) => op,
            None => { offset += 2; continue; }
        };
        let operand_start = offset + 2;
        let next = match op.operand_format() {
            OperandFormat::None => operand_start,
            OperandFormat::U8 => operand_start + 1,
            OperandFormat::U16 => operand_start + 2,
            OperandFormat::I16 => operand_start + 2,
            OperandFormat::U16_U8 => operand_start + 3,
            OperandFormat::U16_U16 => operand_start + 4,
            OperandFormat::U16_I16 => operand_start + 4,
            OperandFormat::Closure => {
                // u16 + u8 + (u8 count × 2)
                let uv_count = *chunk.code.get(operand_start + 2).unwrap_or(&0) as usize;
                operand_start + 3 + uv_count * 2
            }
            OperandFormat::BrTable => {
                let count = *chunk.code.get(operand_start).unwrap_or(&0) as usize;
                operand_start + 2 + count
            }
            OperandFormat::TryTable => {
                let count = *chunk.code.get(operand_start).unwrap_or(&0) as usize;
                operand_start + 1 + count * 3
            }
            OperandFormat::V128Const => operand_start + 16,
            OperandFormat::Shuffle => operand_start + 16,
        };
        // Rewrite the import index for CALL_IMPORT specifically — other
        // U16_U8 opcodes (if any) don't index into the imports table.
        // Vybe stores u16 operands BIG-endian (see `Chunk::read_u16` in
        // `vybe_bytecode/src/chunk.rs:314`).
        if op == Op::CALL_IMPORT {
            let hi = chunk.code[operand_start] as u16;
            let lo = chunk.code[operand_start + 1] as u16;
            let poly_idx = (hi << 8) | lo;
            if let Some(&user_idx) = remap.get(poly_idx as usize) {
                chunk.code[operand_start] = (user_idx >> 8) as u8;
                chunk.code[operand_start + 1] = user_idx as u8;
            }
        }
        offset = next;
    }
}

thread_local! {
    static IN_POLYFILL_BUILD: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

pub fn is_compiling_polyfill() -> bool {
    IN_POLYFILL_BUILD.with(|c| c.get())
}

fn with_polyfill_guard<R>(f: impl FnOnce() -> R) -> R {
    IN_POLYFILL_BUILD.with(|c| c.set(true));
    let result = f();
    IN_POLYFILL_BUILD.with(|c| c.set(false));
    result
}

/// Build all stdlib chunks. Each chunk registers any `ecma:array.*`
/// imports on the passed `imports` chunk (= user program's
/// `chunks[0]`, the module-level imports section per WASM semantics).
/// Returns the stdlib chunks + their export names, in matching order;
/// caller appends the chunks to its own vec.
pub fn build_stdlib(imports: &mut Chunk) -> StdLib {
    let mut chunks = Vec::new();
    let mut exports = Vec::new();

    chunks.push(build_range(imports));             exports.push("__stdlib_range");
    chunks.push(build_sorted(imports));            exports.push("__stdlib_sorted");
    chunks.push(build_sort_in_place(imports));     exports.push("__stdlib_sort_in_place");
    chunks.push(build_sort_with_comparator(imports)); exports.push("__stdlib_sort_with_comparator");
    chunks.push(build_sort_by_key(imports));       exports.push("__stdlib_sort_by_key");
    chunks.push(build_reversed(imports));          exports.push("__stdlib_reversed");
    chunks.push(build_enumerate(imports));         exports.push("__stdlib_enumerate");
    chunks.push(build_zip(imports));               exports.push("__stdlib_zip");
    chunks.push(build_sum(imports));               exports.push("__stdlib_sum");
    chunks.push(build_min(imports));               exports.push("__stdlib_min");
    chunks.push(build_max(imports));               exports.push("__stdlib_max");
    chunks.push(build_pyany(imports));             exports.push("__stdlib_pyany");
    chunks.push(build_pyall(imports));             exports.push("__stdlib_pyall");
    chunks.push(build_compact(imports));           exports.push("__stdlib_compact");
    chunks.push(build_uniq(imports));              exports.push("__stdlib_uniq");
    chunks.push(build_minmax(imports));            exports.push("__stdlib_minmax");
    chunks.push(build_isempty(imports));           exports.push("__stdlib_isempty");
    chunks.push(build_pymap(imports));             exports.push("__stdlib_pymap");
    chunks.push(build_pyfilter(imports));          exports.push("__stdlib_pyfilter");
    chunks.push(build_pynext(imports));            exports.push("__stdlib_pynext");
    chunks.push(build_rand_choice(imports));       exports.push("__stdlib_rand_choice");
    chunks.push(build_rand_shuffle(imports));      exports.push("__stdlib_rand_shuffle");
    chunks.push(build_rand_sample(imports));       exports.push("__stdlib_rand_sample");
    chunks.push(build_rotate(imports));            exports.push("__stdlib_rotate");
    chunks.push(build_array_copy(imports));        exports.push("__stdlib_array_copy");
    chunks.push(build_pow(imports));               exports.push("__stdlib_pow");
    chunks.push(build_sin(imports));               exports.push("__stdlib_sin");
    chunks.push(build_cos(imports));               exports.push("__stdlib_cos");
    chunks.push(build_tan(imports));               exports.push("__stdlib_tan");
    chunks.push(build_asin(imports));              exports.push("__stdlib_asin");
    chunks.push(build_acos(imports));              exports.push("__stdlib_acos");
    chunks.push(build_atan(imports));              exports.push("__stdlib_atan");
    chunks.push(build_atan2(imports));             exports.push("__stdlib_atan2");
    chunks.push(build_log(imports));               exports.push("__stdlib_log");
    chunks.push(build_log10(imports));             exports.push("__stdlib_log10");
    chunks.push(build_exp(imports));               exports.push("__stdlib_exp");
    chunks.push(build_sinh(imports));              exports.push("__stdlib_sinh");
    chunks.push(build_cosh(imports));              exports.push("__stdlib_cosh");
    chunks.push(build_tanh(imports));              exports.push("__stdlib_tanh");
    chunks.push(build_sign(imports));              exports.push("__stdlib_sign");
    chunks.push(build_clamp(imports));             exports.push("__stdlib_clamp");
    chunks.push(build_to_string(imports));         exports.push("__stdlib_tostring");
    chunks.push(build_string_is_null_or_empty(imports)); exports.push("__stdlib_string_is_null_or_empty");
    chunks.push(build_string_is_null_or_whitespace(imports)); exports.push("__stdlib_string_is_null_or_whitespace");
    chunks.push(build_str_insert(imports));        exports.push("__stdlib_str_insert");
    chunks.push(build_str_remove_start(imports));  exports.push("__stdlib_str_remove_start");
    chunks.push(build_str_remove_range(imports));  exports.push("__stdlib_str_remove_range");
    chunks.push(build_str_count(imports));         exports.push("__stdlib_count");
    chunks.push(build_is_numeric(imports));        exports.push("__stdlib_isnumeric");
    chunks.push(build_splice(imports));            exports.push("__stdlib_splice");
    chunks.push(build_floor(imports));             exports.push("__stdlib_floor");
    chunks.push(build_slice(imports));             exports.push("__stdlib_slice");
    chunks.push(build_keys(imports));              exports.push("__stdlib_keys");
    chunks.push(build_has_property(imports));      exports.push("__stdlib_hasproperty");
    chunks.push(build_assign(imports));            exports.push("__stdlib_assign");
    chunks.push(build_instance_of(imports));       exports.push("__stdlib_instanceof");
    chunks.push(build_delete_property(imports));   exports.push("__stdlib_deleteproperty");
    chunks.push(build_array_from(imports));        exports.push("__stdlib_from");
    chunks.push(build_redim(imports));             exports.push("__stdlib_redim");
    chunks.push(build_slice_step(imports));        exports.push("__stdlib_slicestep");
    chunks.push(build_dyn_mul(imports));           exports.push("__stdlib_dynmul");
    chunks.push(build_concat(imports));            exports.push("__stdlib_concat");
    chunks.push(build_string_raw(imports));        exports.push("__stdlib_string_raw");
    chunks.push(build_fmod(imports));              exports.push("__stdlib_fmod");
    chunks.push(build_array_insert(imports));      exports.push("__stdlib_array_insert");
    chunks.push(build_array_remove_at(imports));   exports.push("__stdlib_array_remove_at");
    chunks.push(build_array_remove_value(imports)); exports.push("__stdlib_array_remove_value");
    chunks.push(build_array_insert_range(imports));  exports.push("__stdlib_array_insert_range");
    chunks.push(build_array_set_range(imports));     exports.push("__stdlib_array_set_range");
    chunks.push(build_array_binary_search(imports)); exports.push("__stdlib_array_binary_search");
    chunks.push(build_array_reverse_range(imports)); exports.push("__stdlib_array_reverse_range");
    chunks.push(build_array_last_index_of(imports)); exports.push("__stdlib_array_last_index_of");
    // ── JS-source polyfills ────────────────────────────────────────
    // Compiled at vybex build time via the generic `build_polyfill`
    // plumbing above. Each is a bytecode chunk identical in shape to
    // the hand-emitted ones — bundles into every program and surfaces
    // as a `__vybe_*` global.
    chunks.push(build_polyfill(
        imports, include_str!("polyfills/sprintf.js"), "js", "sprintf"));
    exports.push("__stdlib_sprintf");
    // Order matters: dir() embeds GLOBAL_GET refs to __vybe_dir_read /
    // __vybe_dir_close, which must be registered before dir() runs. The
    // global registration order is the MAPPINGS order (also driven by
    // these `chunks.push` calls), so push the methods first.
    chunks.push(build_dir_read(imports));            exports.push("__stdlib_dir_read");
    chunks.push(build_dir_close(imports));           exports.push("__stdlib_dir_close");
    chunks.push(build_dir(imports));                 exports.push("__stdlib_dir");
    chunks.push(build_file(imports));                exports.push("__stdlib_file");
    chunks.push(build_filemtime(imports));           exports.push("__stdlib_filemtime");
    chunks.push(build_file_exists(imports));         exports.push("__stdlib_file_exists");
    chunks.push(build_is_file(imports));             exports.push("__stdlib_is_file");
    chunks.push(build_is_dir(imports));              exports.push("__stdlib_is_dir");
    chunks.push(build_filesize(imports));            exports.push("__stdlib_filesize");
    chunks.push(build_unlink(imports));              exports.push("__stdlib_unlink");

    // ── Regex adapters: pattern-first → ECMA str-first ─────────────────
    //
    // Python `re.*` and PHP `preg_*` put the regex pattern FIRST per
    // their stdlib convention. ECMA-262 String.prototype.{match, replace,
    // split, matchAll} put the string FIRST (receiver). These adapter
    // chunks bridge the two: take args in language convention, call into
    // `ecma:regexp.*` with reordered args. Same Layer-3 pattern that
    // `String.Format` → `vybe:string.format` uses for the .NET shape.
    chunks.push(build_regex_replace_pat_first(imports)); exports.push("__stdlib_regex_replace_pat_first");
    chunks.push(build_regex_split_pat_first(imports));   exports.push("__stdlib_regex_split_pat_first");
    chunks.push(build_regex_match_all_pat_first(imports)); exports.push("__stdlib_regex_match_all_pat_first");

    StdLib { chunks, exports }
}

pub struct StdLib {
    pub chunks: Vec<Chunk>,
    pub exports: Vec<&'static str>,
}

impl StdLib {
    pub fn get(&self, name: &str) -> Option<usize> {
        self.exports.iter().position(|&n| n == name)
    }
}

// ── range(start, stop, step) → array ────────────────────────
// Every dynamic-array op routes through `common::collections::emit_*`
// so the emitted bytecode imports `ecma:array.*` — works natively on
// v8, on Vybe (registered handlers), and on plain wasmtime with the
// polyfill module. Raw ARRAY_* opcodes are Vybe-only and have been
// removed.
fn build_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_range");
    c.arity = 3; // start, stop, step
    c.local_count = 4;
    let start = 0u16;
    let stop = 1;
    let step = 2;
    let result = 3;

    // result = []
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, stop, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);

    // result.push(i) — push returns new_length (ECMA-262); drop it.
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, start, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sorted(array) → array (insertion sort — O(n²) but works) ──
fn build_sorted(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sorted");
    c.arity = 1;
    c.local_count = 6; // arr(0) + result(1) + i(2) + j(3) + len(4) + key(5)
    let arr = 0u16;
    let result = 1;
    let i = 2;
    let j = 3;
    let len = 4;
    let key = 5;

    // Copy input array → result (so we don't mutate the original)
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::CONST, max, 0);
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = result.length
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // Insertion sort: for i = 1 to len-1
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = result[i]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    // while j >= 0 && result[j] > key
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // result[j+1] = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    // Now stack: [result, j+1] — need value = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0); c.patch_loop(inner_loop_p);
    c.emit_end(0); c.patch_block(inner_block_p);

    // result[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0); c.patch_loop(outer_loop_p);
    c.emit_end(0); c.patch_block(outer_block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sort_in_place(array) → same array, mutated ──────────────
// In-place insertion sort. Used by every language whose surface syntax for
// sorting is in-place: C# `list.Sort()`, VB `list.Sort()`, JS `arr.sort()`,
// Python `list.sort()`, Pascal `Sort(arr)`. The walker normalizes each form
// into a canonical builtin call which routes here through compiler_common.
//
// Insertion sort is O(n²) but small and works on arbitrary value comparisons
// via dyn_gt. Higher-perf algorithms can be added behind the same name later.
fn build_sort_in_place(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sort_in_place");
    c.arity = 1;
    c.local_count = 5; // arr(0) + i(1) + j(2) + len(3) + key(4)
    let arr = 0u16;
    let i = 1;
    let j = 2;
    let len = 3;
    let key = 4;

    // len = arr.length
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 1
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    // while j >= 0 && arr[j] > key
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0); c.patch_loop(inner_loop_p);
    c.emit_end(0); c.patch_block(inner_block_p);

    // arr[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0); c.patch_loop(outer_loop_p);
    c.emit_end(0); c.patch_block(outer_block_p);

    // return arr (same reference, now sorted in place)
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sort_with_comparator(array, fn) → same array, sorted using fn ──
// Same insertion sort as sort_in_place, but uses `fn(a, b)` for
// comparison instead of `dyn_gt`. The comparator returns:
//   negative → a before b (no swap)
//   zero     → equal (no swap)
//   positive → b before a (swap)
// This is the standard JS `Array.sort(compareFn)` contract.
fn build_sort_with_comparator(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sort_with_comparator");
    c.arity = 2;
    c.local_count = 6; // arr(0) + cmp(1) + i(2) + j(3) + len(4) + key(5)
    let arr = 0u16;
    let cmp = 1;
    let i = 2;
    let j = 3;
    let len = 4;
    let key = 5;

    // len = arr.length
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 1
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    // while j >= 0 && cmp(arr[j], key) > 0
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop

    // call cmp(arr[j], key) → result
    c.emit_op_u16(Op::LOCAL_GET, cmp, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op_u8(Op::CALL_REF, 2, 0);
    // result > 0 → swap needed
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0); c.patch_loop(inner_loop_p);
    c.emit_end(0); c.patch_block(inner_block_p);

    // arr[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0); c.patch_loop(outer_loop_p);
    c.emit_end(0); c.patch_block(outer_block_p);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── reversed(array) → array ─────────────────────────────────
fn build_reversed(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_reversed");
    c.arity = 1;
    c.local_count = 3; // arr(0) + result(1) + i(2)
    let arr = 0u16;
    let result = 1;
    let i = 2;

    // result = []
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // i = arr.length - 1
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── enumerate(array) → [[0,a],[1,b],...] ────────────────────
fn build_enumerate(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_enumerate");
    c.arity = 1;
    c.local_count = 4; // arr(0) + result(1) + i(2) + len(3)
    let arr = 0u16;
    let result = 1;
    let i = 2;
    let len = 3;

    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    // Build pair [i, arr[i]], then push onto result.
    // array_push takes [array, value] — so emit result first, then pair.
    c.emit_op_u16(Op::LOCAL_GET, result, 0); // result on stack
    c.emit_op_u16(Op::LOCAL_GET, i, 0);      // i
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);             // arr[i]
    crate::emitter::collections::emit_array_pair_into(imports, &mut c, 0);      // pair = [i, arr[i]]
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);            // result.push(pair)
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── zip(a, b) → [[a0,b0],[a1,b1],...] ──────────────────────
fn build_zip(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_zip");
    c.arity = 2;
    c.local_count = 5; // a(0) + b(1) + result(2) + i(3) + len(4)
    let a = 0u16;
    let b = 1;
    let result = 2;
    let i = 3;
    let len = 4;

    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = min(a.length, b.length) — use a.length for simplicity
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    // result.push([a[i], b[i]])
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, b, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_array_pair_into(imports, &mut c, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sum(array) → number ─────────────────────────────────────
fn build_sum(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sum");
    c.arity = 1;
    c.local_count = 4; // arr(0) + total(1) + i(2) + len(3)
    let arr = 0u16;
    let total = 1;
    let i = 2;
    let len = 3;

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, total, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, total, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── any(iter) / all(iter) → bool ─────────────────────────────
// Python `any(iter)` / `all(iter)` — bare-iterable shape (no callback).
// Spec-shape equivalent of `arr.some(Boolean)` / `arr.every(Boolean)`
// without requiring callers to materialize the Boolean fn ref. Mirrors
// the polymorphic ARRAY_GET semantics so it works on Array, Map, and
// String operands transparently.
fn build_pyany(imports: &mut Chunk) -> Chunk {
    build_any_all(imports, "__stdlib_pyany", true)
}

fn build_pyall(imports: &mut Chunk) -> Chunk {
    build_any_all(imports, "__stdlib_pyall", false)
}

fn build_any_all(imports: &mut Chunk, name: &str, is_any: bool) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = 1;
    c.local_count = 3; // arr(0) + i(1) + len(2)
    let arr = 0u16;
    let i = 1;
    let len = 2;

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0); c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop → fell through

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::DYN_TO_BOOL, 0);
    if is_any {
        // any: if truthy → return true
        let to_continue = c.emit_jump(Op::BR_IF_FALSE, 0);
        c.emit_op(Op::TRUE, 0);
        c.emit_op(Op::RETURN, 0);
        c.patch_jump(to_continue);
    } else {
        // all: if falsy → return false
        let to_continue = c.emit_jump(Op::BR_IF_TRUE, 0);
        c.emit_op(Op::FALSE, 0);
        c.emit_op(Op::RETURN, 0);
        c.patch_jump(to_continue);
    }

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    // Loop fell through: any → false, all → true
    if is_any { c.emit_op(Op::FALSE, 0); } else { c.emit_op(Op::TRUE, 0); }
    c.emit_op(Op::RETURN, 0);
    c
}

// ── min(array) → value ──────────────────────────────────────
fn build_min(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_min");
    c.arity = 1;
    c.local_count = 4; // arr(0) + best(1) + i(2) + len(3)
    let arr = 0u16;
    let best = 1;
    let i = 2;
    let len = 3;

    // best = arr[0]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    // if arr[i] < best: best = arr[i]
    // block must wrap ALL condition operands + comparison + body
    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if NOT less than
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(skip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── max(array) → value ──────────────────────────────────────
fn build_max(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_max");
    c.arity = 1;
    c.local_count = 4;
    let arr = 0u16;
    let best = 1;
    let i = 2;
    let len = 3;

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if NOT greater than
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(skip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── compact(arr) → arr without nulls (Ruby Array#compact) ──
fn build_compact(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_compact");
    c.arity = 1;
    c.local_count = 5; // arr(0) + result(1) + i(2) + len(3) + elem(4)
    let arr = 0u16;
    let result = 1;
    let i = 2;
    let len = 3;
    let elem = 4;

    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0); c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0); c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);

    // elem = arr[i]; stash into local
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, elem, 0); c.emit_op(Op::DROP, 0);

    // if !is_null(elem) → result.push(elem)
    c.emit_op_u16(Op::LOCAL_GET, elem, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let skip = c.emit_jump(Op::BR_IF_TRUE, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, elem, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.patch_jump(skip);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── isEmpty(arr) → bool (Ruby Array#empty?) ──
fn build_isempty(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isempty");
    c.arity = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── minmax(arr) → [min, max] (Ruby Array#minmax) ──
fn build_minmax(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_minmax");
    c.arity = 1;
    c.local_count = 3; // arr(0) + min(1) + max(2)
    let arr = 0u16;
    let min_g = 1;
    let max_g = 2;

    // min = __vybe_min(arr); max = __vybe_max(arr); return [min, max]
    let name_min = c.add_constant(Value::String(Arc::from("__vybe_min")));
    c.emit_op_u16(Op::GLOBAL_GET, name_min, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, min_g, 0); c.emit_op(Op::DROP, 0);

    let name_max = c.add_constant(Value::String(Arc::from("__vybe_max")));
    c.emit_op_u16(Op::GLOBAL_GET, name_max, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, max_g, 0); c.emit_op(Op::DROP, 0);

    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op(Op::DUP, 0);
    c.emit_op_u16(Op::LOCAL_GET, min_g, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0); c.emit_op(Op::DROP, 0);
    c.emit_op(Op::DUP, 0);
    c.emit_op_u16(Op::LOCAL_GET, max_g, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0); c.emit_op(Op::DROP, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── uniq(arr) → arr with duplicates removed (Ruby Array#uniq) ──
fn build_uniq(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_uniq");
    c.arity = 1;
    c.local_count = 5; // arr(0) + result(1) + i(2) + len(3) + elem(4)
    let arr = 0u16;
    let result = 1;
    let i = 2;
    let len = 3;
    let elem = 4;

    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0); c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0); c.emit_op(Op::DROP, 0);

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);

    // elem = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, elem, 0); c.emit_op(Op::DROP, 0);

    // if !result.includes(elem) result.push(elem)
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, elem, 0);
    let inc_idx = imports.add_import("ecma:array", "includes");
    c.emit_op_u16(Op::CALL_IMPORT, inc_idx, 0); c.emit(2u8, 0);
    let already = c.emit_jump(Op::BR_IF_TRUE, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, elem, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0); c.emit_op(Op::DROP, 0);
    c.patch_jump(already);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── pymap(fn, iter) — Python `map(fn, iter)` shape adapter ──
// Wraps ECMA `Array.prototype.map(fn)` (§23.1.3.21) with swapped
// args: Python passes (fn, iter), ECMA expects (iter, fn).
fn build_pymap(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pymap");
    c.arity = 2; // fn(0), iter(1)
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // iter
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // fn
    let idx = imports.add_import("ecma:array", "map");
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0); c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── pyfilter(fn, iter) — Python `filter(fn, iter)` shape adapter ──
fn build_pyfilter(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pyfilter");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let idx = imports.add_import("ecma:array", "filter");
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0); c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── pynext(iter, default?) — Python `next(iter, default)` ──
// Returns and removes the first element. Default returned when empty.
fn build_pynext(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pynext");
    c.arity = 2;
    c.local_count = 2;
    // if iter.length == 0 → return default (or null when default omitted)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_EQ, 0);
    let to_consume = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // default
    c.emit_op(Op::RETURN, 0);
    c.patch_jump(to_consume);
    // shift first element off iter
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let sh_idx = imports.add_import("ecma:array", "shift");
    c.emit_op_u16(Op::CALL_IMPORT, sh_idx, 0); c.emit(1u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── rand_choice(arr) — random element via ecma:math.random ──
// `arr[Math.floor(Math.random() * arr.length)]`. Returns null on empty.
fn build_rand_choice(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_rand_choice");
    c.arity = 1;
    c.local_count = 3; // arr(0), len(1), idx(2)
    let arr = 0u16;
    let len = 1;
    let idx = 2;
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0); c.emit_op(Op::DROP, 0);

    // empty? return null
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_EQ, 0);
    let not_empty = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op(Op::NULL, 0);
    c.emit_op(Op::RETURN, 0);
    c.patch_jump(not_empty);

    // idx = floor(random() * len)
    let r_idx = imports.add_import("ecma:math", "random");
    c.emit_op_u16(Op::CALL_IMPORT, r_idx, 0); c.emit(0u8, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::F64_FROM_I32, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    c.emit_op_u16(Op::LOCAL_SET, idx, 0); c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── rand_shuffle(arr) — in-place Fisher-Yates with ecma:math.random ──
fn build_rand_shuffle(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_rand_shuffle");
    c.arity = 1;
    c.local_count = 5; // arr(0), i(1), j(2), tmp(3), len(4)
    let arr = 0u16;
    let i = 1;
    let j = 2;
    let tmp = 3;
    let len = 4;

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0); c.emit_op(Op::DROP, 0);

    // i = len - 1
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    // while i > 0
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_LE, 0);
    c.emit_br_if(1, 0); // exit if i <= 0

    // j = floor(random() * (i + 1))
    let r_idx = imports.add_import("ecma:math", "random");
    c.emit_op_u16(Op::CALL_IMPORT, r_idx, 0); c.emit(0u8, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op(Op::F64_FROM_I32, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0); c.emit_op(Op::DROP, 0);

    // tmp = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, tmp, 0); c.emit_op(Op::DROP, 0);
    // arr[i] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::ARRAY_SET, 0); c.emit_op(Op::DROP, 0);
    // arr[j] = tmp
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op_u16(Op::LOCAL_GET, tmp, 0);
    c.emit_op(Op::ARRAY_SET, 0); c.emit_op(Op::DROP, 0);

    // i--
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── rand_sample(arr, k) — Python `random.sample(seq, k)` ──
// Returns a new array of k elements without replacement. Uses
// shuffle then slice — O(n) memory, simple and correct.
fn build_rand_sample(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_rand_sample");
    c.arity = 2;
    c.local_count = 3;
    // copy = arr.slice(0, len) — duplicate so shuffle doesn't mutate caller
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    let sl_idx = imports.add_import("ecma:array", "slice");
    c.emit_op_u16(Op::CALL_IMPORT, sl_idx, 0); c.emit(3u8, 0);
    c.emit_op_u16(Op::LOCAL_SET, 2, 0); c.emit_op(Op::DROP, 0);

    // shuffle copy in place
    let sh_name = c.add_constant(Value::String(Arc::from("__vybe_rand_shuffle")));
    c.emit_op_u16(Op::GLOBAL_GET, sh_name, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    c.emit_op(Op::DROP, 0);

    // return copy.slice(0, k)
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::CALL_IMPORT, sl_idx, 0); c.emit(3u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── rotate(arr, n) — Ruby `Array#rotate(n)` ──
// Returns new array rotated n positions left. n defaults to 1; negative
// rotates right. Implemented as `arr.slice(n).concat(arr.slice(0, n))`
// after normalizing n into [0, len).
fn build_rotate(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_rotate");
    c.arity = 2;
    c.local_count = 4; // arr(0), n(1), len(2), n_norm(3)
    let arr = 0u16;
    let n = 1;
    let len = 2;
    let n_norm = 3;

    // n defaults to 1 if null/undefined
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let n_ok = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, n, 0); c.emit_op(Op::DROP, 0);
    c.patch_jump(n_ok);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0); c.emit_op(Op::DROP, 0);

    // n_norm = ((n % len) + len) % len  — handles negative n
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    let fmod_name = c.add_constant(Value::String(Arc::from("__vybe_fmod")));
    c.emit_op_u16(Op::GLOBAL_GET, fmod_name, 0);
    // fmod expects [a, b] before func ref; CALL_REF expects [func, args]
    // We have [n, len, fmod]. Stash + reload.
    // Simpler: just inline the modulo via DYN_MOD if it exists. Otherwise:
    // For now, assume n < len and n >= -len: fix via emit
    c.emit_op(Op::DROP, 0); // drop fmod ref, redo cleanly
    c.emit_op(Op::DROP, 0); // drop len
    c.emit_op(Op::DROP, 0); // drop n
    // Recompute properly: stack []
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::I32_REM_S, 0);          // n % len
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_ADD, 0);           // + len
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::I32_REM_S, 0);           // % len → n_norm
    c.emit_op_u16(Op::LOCAL_SET, n_norm, 0); c.emit_op(Op::DROP, 0);

    // result = arr.slice(n_norm, len).concat(arr.slice(0, n_norm))
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, n_norm, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    let sl_idx = imports.add_import("ecma:array", "slice");
    c.emit_op_u16(Op::CALL_IMPORT, sl_idx, 0); c.emit(3u8, 0);
    // [first_part]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, n_norm, 0);
    c.emit_op_u16(Op::CALL_IMPORT, sl_idx, 0); c.emit(3u8, 0);
    // [first_part, second_part]
    let cc_idx = imports.add_import("ecma:array", "concat");
    c.emit_op_u16(Op::CALL_IMPORT, cc_idx, 0); c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_copy(src, dst, count) — C# `Array.Copy(src, dst, count)` ──
// Per .NET spec: copies `count` elements from src[0..] to dst[0..].
fn build_array_copy(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_copy");
    c.arity = 3;
    c.local_count = 4; // src(0), dst(1), count(2), i(3)
    let src = 0u16;
    let dst = 1;
    let count = 2;
    let i = 3;

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);

    // dst[i] = src[i]
    c.emit_op_u16(Op::LOCAL_GET, dst, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::ARRAY_SET, 0); c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    // .NET Array.Copy returns void
    c.emit_op(Op::NULL, 0);
    c.emit_op(Op::RETURN, 0);
    let _ = imports; // silence unused
    c
}

// ── pow(base, exp) → number (integer exponent by repeated mul) ──
fn build_pow(_imports: &mut Chunk) -> Chunk {
    // Bytecode-only fallback for `pow(base, exp)`. Handles INTEGER exponents
    // (positive, zero, negative) using a multiply loop. Fractional exponents
    // require floating-point exp/log which WASM doesn't have as standard
    // opcodes — Vybe overrides `__vybe_pow` with a native f64.powf at runtime
    // (polyfill pattern), so this fallback only runs on non-Vybe runtimes
    // and only needs to be correct for the common integer-exp case.
    let mut c = Chunk::new("__stdlib_pow");
    c.arity = 2;
    c.local_count = 4; // base(0) + exp(1) + result(2) + n(3)
    let base = 0u16;
    let exp = 1;
    let result = 2;
    let n = 3;

    // n = abs(exp) — branchless via select would need both values; use a flag
    // We compute n = (exp < 0) ? -exp : exp
    c.emit_op_u16(Op::LOCAL_GET, exp, 0);
    c.emit_op_u16(Op::LOCAL_SET, n, 0);
    c.emit_op(Op::DROP, 0);
    // if n < 0 then n = -n
    let pos_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip negate if NOT negative
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op(Op::F64_NEG, 0);
    c.emit_op_u16(Op::LOCAL_SET, n, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(pos_block_p);

    // result = 1.0
    let one = c.add_constant(Value::F64(1.0));
    c.emit_op_u16(Op::CONST, one, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // while n > 0: result *= base; n -= 1
    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, base, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, n, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    // If original exp was negative, take reciprocal: result = 1.0 / result
    let recip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, exp, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip reciprocal if NOT negative
    c.emit_op_u16(Op::CONST, one, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::F64_DIV, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(recip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_unary_env_math(imports: &mut Chunk, chunk_name: &str, env_name: &str) -> Chunk {
    let idx = imports.add_import("env", env_name);
    let mut c = Chunk::new(chunk_name);
    c.arity = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(1, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_binary_env_math(imports: &mut Chunk, chunk_name: &str, env_name: &str) -> Chunk {
    let idx = imports.add_import("env", env_name);
    let mut c = Chunk::new(chunk_name);
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_sin(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_sin", "sin")
}

fn build_cos(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_cos", "cos")
}

fn build_tan(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_tan", "tan")
}

fn build_asin(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_asin", "asin")
}

fn build_acos(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_acos", "acos")
}

fn build_atan(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_atan", "atan")
}

fn build_atan2(imports: &mut Chunk) -> Chunk {
    build_binary_env_math(imports, "__stdlib_atan2", "atan2")
}

fn build_sinh(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_sinh", "sinh")
}

fn build_cosh(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_cosh", "cosh")
}

fn build_tanh(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_tanh", "tanh")
}

fn build_log(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_log", "log")
}

fn build_log10(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_log10", "log10")
}

fn build_exp(imports: &mut Chunk) -> Chunk {
    build_unary_env_math(imports, "__stdlib_exp", "exp")
}

fn build_sign(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sign");
    c.arity = 1;
    c.local_count = 1;
    let value = 0u16;
    let one = c.add_constant(Value::F64(1.0));
    let zero = c.add_constant(Value::F64(0.0));
    let minus_one = c.add_constant(Value::F64(-1.0));

    let skip_positive = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::CONST, one, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(skip_positive);

    let skip_negative = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::CONST, minus_one, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(skip_negative);

    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_clamp(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_clamp");
    c.arity = 3;
    c.local_count = 3;
    let value = 0u16;
    let min = 1u16;
    let max = 2u16;

    let skip_min = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::LOCAL_GET, min, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, min, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(skip_min);

    let skip_max = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::LOCAL_GET, max, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, max, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(skip_max);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── toString(value) → string ────────────────────────────────
// "" + value triggers dyn_add string coercion in the VM
fn build_to_string(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_tostring");
    c.arity = 1;
    c.local_count = 1;
    let val = 0u16;
    let empty = c.add_constant(Value::String(std::sync::Arc::from("")));
    c.emit_op_u16(Op::CONST, empty, 0);
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── string_is_null_or_empty(value) → bool ─────────────────
fn build_string_is_null_or_empty(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_string_is_null_or_empty");
    c.arity = 1;
    c.local_count = 1;
    let value = 0u16;

    let non_null = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op(Op::TRUE, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(non_null);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── string_is_null_or_whitespace(value) → bool ─────────────
fn build_string_is_null_or_whitespace(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_string_is_null_or_whitespace");
    c.arity = 1;
    c.local_count = 1;
    let value = 0u16;

    let non_null = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op(Op::TRUE, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(non_null);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op(Op::STR_TRIM, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── str_insert(str, index, value) → string ────────────────
fn build_str_insert(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_str_insert");
    c.arity = 3;
    c.local_count = 3;
    let value = 2u16;
    let max = c.add_constant(Value::I32(i32::MAX));

    // prefix = str[0:index]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);

    // prefix + value (keeps current coercion behavior for non-string values)
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op(Op::DYN_ADD, 0);

    // + suffix = str[index:]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::CONST, max, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── str_remove_start(str, start) → string ─────────────────
fn build_str_remove_start(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_str_remove_start");
    c.arity = 2;
    c.local_count = 2;

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── str_remove_range(str, start, count) → string ──────────
fn build_str_remove_range(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_str_remove_range");
    c.arity = 3;
    c.local_count = 3;
    let max = c.add_constant(Value::I32(i32::MAX));

    // prefix = str[0:start]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);

    // suffix = str[start+count:]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::CONST, max, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── count(haystack, needle) → int ───────────────────────────
// Count non-overlapping occurrences using substring + indexOf loop
fn build_str_count(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_count");
    c.arity = 2;
    c.local_count = 4;
    let haystack = 0u16;
    let needle = 1;
    let count = 2;
    let pos = 3;

    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, count, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, pos, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, haystack, 0);
    c.emit_op_u16(Op::LOCAL_GET, pos, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::CONST, max, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op_u16(Op::LOCAL_GET, needle, 0);
    c.emit_op(Op::STR_INDEX_OF, 0);
    // Save indexOf result to local (don't use DUP — value can't cross block boundary)
    let idx_result = 4u16; // reuse local slot (local_count=4, slot 4 is beyond declared but safe with extra locals)
    c.local_count = 5; // need one more local for idx_result
    c.emit_op_u16(Op::LOCAL_SET, idx_result, 0);
    c.emit_op(Op::DROP, 0);
    // Check if index < 0
    c.emit_op_u16(Op::LOCAL_GET, idx_result, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_br_if(1, 0); // exit loop if index < 0
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, count, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, pos, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, pos, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── splice(arr, index, deleteCount) → removed_elements ──────
// Returns array of removed elements. Mutates arr by removing elements.
// Pure bytecode: build new array from arr[0:index] + arr[index+deleteCount:end]
fn build_splice(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_splice");
    c.arity = 3;
    c.local_count = 6; // arr(0) + index(1) + delete_count(2) + result(3) + i(4) + end(5)
    let arr = 0u16;
    let index = 1;
    let delete_count = 2;
    let result_local = 3;
    let i = 4;
    let end = 5;

    // result = [] (removed elements)
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result_local, 0);
    c.emit_op(Op::DROP, 0);

    // Collect removed elements: arr[index..index+deleteCount]
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_GET, delete_count, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, result_local, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    // Return removed elements (actual array mutation would need more complex bytecode)
    c.emit_op_u16(Op::LOCAL_GET, result_local, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── isNumeric(value) → bool ─────────────────────────────────
// Check if value is a number type using ref_typeof opcode.
fn build_is_numeric(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isnumeric");
    c.arity = 1;
    c.local_count = 1; // val(0)
    let val = 0u16;

    // Check if type is "number" (covers I32, I64, F64)
    // Block must wrap ALL values consumed inside it (typeof result + STR_EQUALS + DUP)
    let num_str = c.add_constant(Value::String(std::sync::Arc::from("number")));
    let done_block_p = c.emit_block(0);
    // typeof(val) → string
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, num_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);

    // If true, save and skip second check
    let result_slot = 1u16;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_SET, result_slot, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, result_slot, 0);
    c.emit_br_if(0, 0); // skip to end if already true

    // Also check if typeof is "i32"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    let i32_str = c.add_constant(Value::String(std::sync::Arc::from("i32")));
    c.emit_op_u16(Op::CONST, i32_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op_u16(Op::LOCAL_SET, result_slot, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_end(0); c.patch_block(done_block_p);
    c.emit_op_u16(Op::LOCAL_GET, result_slot, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── floor(n) → int — wraps f64_floor opcode ────────────────
fn build_floor(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_floor");
    c.arity = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::F64_FLOOR, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── slice(arr, start, end) → array — wraps array_slice opcode
fn build_slice(imports: &mut Chunk) -> Chunk {
    // Polymorphic slice: handles BOTH strings and arrays. The walker doesn't
    // know whether `obj[1..3]` operates on a string or an array, so the
    // canonical slice helper does a runtime type check via `ref_is_string`
    // and dispatches to `str_substring` or `array_slice` accordingly.
    //
    // Used by every language whose surface syntax for slicing is `[start..end]`:
    // C# `arr[1..3]` / `s[0..5]`, Python `arr[1:3]` / `s[0:5]`, etc.
    let mut c = Chunk::new("__stdlib_slice");
    c.arity = 3;
    c.local_count = 3; // obj + start + end
    let obj = 0u16;
    let start = 1u16;
    let end = 2u16;

    // if ref_is_string(obj) → str_substring; else → array_slice
    let str_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, obj, 0);
    c.emit_op(Op::REF_IS_STRING, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip string branch if NOT string

    // String branch: [obj, start, end] → str_substring
    c.emit_op_u16(Op::LOCAL_GET, obj, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op(Op::RETURN, 0);

    // Array branch
    c.emit_end(0); c.patch_block(str_block_p);
    c.emit_op_u16(Op::LOCAL_GET, obj, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── keys(obj) → array of string keys ────────────────────────
// Iterates object properties, collects non-internal keys.
fn build_keys(imports: &mut Chunk) -> Chunk {
    // Can't iterate properties in pure bytecode without host support.
    // Use dict_keys host call pattern — but that's what we're trying to avoid.
    // Fallback: return empty array. On Vybe, host fn handles it.
    let mut c = Chunk::new("__stdlib_keys");
    c.arity = 1;
    c.local_count = 1;
    // Return empty array as fallback (properties aren't enumerable in pure WASM)
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── hasProperty(obj, key) → bool ────────────────────────────
fn build_has_property(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_hasproperty");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // obj
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // key
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── assign(target, source) → target with source props merged ─
fn build_assign(_imports: &mut Chunk) -> Chunk {
    // Can't iterate source properties in pure bytecode.
    // Fallback: return target unchanged.
    let mut c = Chunk::new("__stdlib_assign");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // return target
    c.emit_op(Op::RETURN, 0);
    c
}

// ── instanceOf(obj, type_name) → bool ───────────────────────
fn build_instance_of(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_instanceof");
    c.arity = 2;
    c.local_count = 2;
    // ref_test needs a constant pool string, but we have a dynamic value.
    // Workaround: compare __type property with the type name string.
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // obj
    let type_key = c.add_constant(Value::String(std::sync::Arc::from("__type")));
    c.emit_op_u16(Op::CONST, type_key, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0); // obj["__type"]
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // type_name
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── deleteProperty(obj, key) → bool ─────────────────────────
fn build_delete_property(imports: &mut Chunk) -> Chunk {
    // Can't delete properties in pure bytecode. Set to null as fallback.
    let mut c = Chunk::new("__stdlib_deleteproperty");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // obj
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // key
    c.emit_op(Op::NULL, 0);             // value = null
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op(Op::TRUE, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── from(iterable) → array copy ─────────────────────────────
fn build_array_from(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_from");
    c.arity = 1;
    c.local_count = 1;
    // Slice the entire array (copy)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    c.emit_op_u16(Op::CONST, max, 0);
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── redim(arr, new_size) → resized array ────────────────────
fn build_redim(imports: &mut Chunk) -> Chunk {
    // Create new array of new_size, copy elements from old
    let mut c = Chunk::new("__stdlib_redim");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // arr
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // new_size
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sliceStep(arr, start, end, step) → array ─────────────────
fn build_slice_step(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_slicestep");
    c.arity = 4;
    c.local_count = 7; // arr(0) start(1) end(2) step(3) result(4) i(5) cond(6)
    let zero = c.add_constant(Value::I32(0));

    // result = new array
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 4, 0);
    c.emit_op(Op::DROP, 0);
    // i = start
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, 5, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // Compute condition: if step > 0 then i < end else i > end
    // Store in local 6 (cond) to avoid value-on-stack across branches
    let pos_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip positive branch if step <= 0
    // positive step: cond = i < end
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op_u16(Op::LOCAL_SET, 6, 0);
    c.emit_op(Op::DROP, 0);
    let skip_neg_p = c.emit_block(0);
    c.emit_br(1, 0); // skip negative branch (jump past skip_neg block end + neg block)
    c.emit_end(0); c.patch_block(skip_neg_p);
    c.emit_end(0); c.patch_block(pos_block_p);

    // negative step: cond = i > end
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op_u16(Op::LOCAL_SET, 6, 0);
    c.emit_op(Op::DROP, 0);

    // Check condition — exit if false
    c.emit_op_u16(Op::LOCAL_GET, 6, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop (depth 1 = outer block)

    // bounds check: skip push if i < 0 or i >= arr.length
    // Block must wrap the condition values consumed by br_if inside it
    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_br_if(0, 0); // skip push if i < 0
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_br_if(0, 0); // skip push if i >= length
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(skip_block_p);

    // i = i + step
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 5, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── dynMul(a, b) → string repeat or numeric multiply ─────────
fn build_dyn_mul(_imports: &mut Chunk) -> Chunk {
    use std::sync::Arc;
    let mut c = Chunk::new("__stdlib_dynmul");
    c.arity = 2;
    c.local_count = 2;
    let str_tag = c.add_constant(Value::String(Arc::from("string")));
    // if typeof(a) == "string": return str_repeat(a, b)
    let a_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, str_tag, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if a is NOT string
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::STR_REPEAT, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(a_block_p);
    // if typeof(b) == "string": return str_repeat(b, a)
    let b_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, str_tag, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if b is NOT string
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::STR_REPEAT, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(b_block_p);
    // numeric
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── sort_by_key(array, keyFn) → same array, sorted by keyFn(x) ──
// .NET LINQ OrderBy(keySelector): insertion sort where comparisons use
// keyFn(a) vs keyFn(b) instead of a vs b directly. The keyFn is a
// 1-arg function that extracts the sort key from each element.
// `OrderBy(x => x)` is identity (plain sort). `OrderBy(x => x.name)`
// sorts by the name property.
fn build_sort_by_key(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_sort_by_key");
    c.arity = 2;
    c.local_count = 7; // arr(0) + keyFn(1) + i(2) + j(3) + len(4) + key(5) + keyVal(6)
    let arr = 0u16;
    let key_fn = 1;
    let i = 2;
    let j = 3;
    let len = 4;
    let key = 5;
    let key_val = 6;

    // len = arr.length
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 1
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    // keyVal = keyFn(key)
    c.emit_op_u16(Op::LOCAL_GET, key_fn, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, key_val, 0);
    c.emit_op(Op::DROP, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    // while j >= 0 && keyFn(arr[j]) > keyVal
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop

    // compare: keyFn(arr[j]) > keyVal
    c.emit_op_u16(Op::LOCAL_GET, key_fn, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, key_val, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0); c.patch_loop(inner_loop_p);
    c.emit_end(0); c.patch_block(inner_block_p);

    // arr[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::emitter::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0); c.patch_loop(outer_loop_p);
    c.emit_end(0); c.patch_block(outer_block_p);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── concat(a, b) → polymorphic concat ───────────────────────
// If `a` is a string, do str_concat. If `a` is an array, do array_concat.
// Runtime dispatch using ref_is_string. Pure WASM bytecode.
fn build_concat(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_concat");
    c.arity = 2; // a, b
    c.local_count = 2; // a(0) + b(1)
    let a = 0u16;
    let b = 1u16;

    // if ref_is_string(a) → str_concat
    let str_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op(Op::REF_IS_STRING, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip string path if NOT string

    // String path: str_concat(a, b)
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op_u16(Op::LOCAL_GET, b, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op(Op::RETURN, 0);

    c.emit_end(0); c.patch_block(str_block_p);

    // Array path: array_concat(a, b)
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op_u16(Op::LOCAL_GET, b, 0);
    crate::emitter::collections::emit_concat_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);

    c
}

// ── String.raw(strings, ...values) → interleave strings and values ──
// Tagged template function that returns the raw string without escape processing.
// strings[0] + values[0] + strings[1] + values[1] + ... + strings[N]
// Since this is called as a tagged template, strings is an array and
// values are individual args. With rest params, values is already an array.
fn build_string_raw(imports: &mut Chunk) -> Chunk {
    use std::sync::Arc;

    let mut c = Chunk::new("__stdlib_string_raw");
    c.arity = 2; // strings_array, values_array (rest-packed by caller)
    c.local_count = 5; // strings(0) + values(1) + result(2) + i(3) + len(4)
    let strings = 0u16;
    let values = 1u16;
    let result = 2u16;
    let i = 3u16;
    let len = 4u16;

    // result = ""
    let empty = c.add_constant(Value::String(Arc::from("")));
    c.emit_op_u16(Op::CONST, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = strings.length
    c.emit_op_u16(Op::LOCAL_GET, strings, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 0
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    // loop: while i < len
    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit loop

    // result += strings[i]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, strings, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // if i < values.length: result += String(values[i])
    let skip_val_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, values, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // skip if i >= values.length

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, values, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::STR_CONCAT, 0); // dyn_add would also work since result is string
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_end(0); c.patch_block(skip_val_p);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── fmod(a, b) → a % b (floating-point remainder) ──────────
// WASM has no f64.rem. Pure bytecode: a - trunc(a/b) * b.
// Host can override __vybe_fmod with native fmod for performance.
fn build_fmod(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_fmod");
    c.arity = 2; // a, b
    c.local_count = 2; // a(0) + b(1)
    let a = 0u16;
    let b = 1u16;

    // result = a - trunc(a / b) * b
    c.emit_op_u16(Op::LOCAL_GET, a, 0);   // a
    c.emit_op_u16(Op::LOCAL_GET, a, 0);   // a
    c.emit_op_u16(Op::LOCAL_GET, b, 0);   // b
    c.emit_op(Op::F64_DIV, 0);            // a / b
    c.emit_op(Op::F64_TRUNC, 0);          // trunc(a / b)
    c.emit_op_u16(Op::LOCAL_GET, b, 0);   // b
    c.emit_op(Op::F64_MUL, 0);            // trunc(a / b) * b
    c.emit_op(Op::F64_SUB, 0);            // a - trunc(a / b) * b
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_insert(arr, index, value) → null ──────────────────────────────
// splice(arr, index, 0, value) — inserts value at index without removing anything.
fn build_array_insert(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_insert");
    c.arity = 3; // arr, index, value
    c.local_count = 3;
    let arr = 0u16;
    let index = 1;
    let value = 2;

    // splice(arr, index, 0, value)
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op(Op::I32_CONST_0, 0); // deleteCount = 0
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    let splice = imports.add_import("ecma:array", "splice");
    c.emit_op_u16(Op::CALL_IMPORT, splice, 0);
    c.emit(4u8, 0); // 4 args
    c.emit_op(Op::DROP, 0); // drop returned removed-elements array
    c.emit_op(Op::NULL, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_remove_at(arr, index) → null ──────────────────────────────────
// splice(arr, index, 1) — removes 1 element at index.
fn build_array_remove_at(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_remove_at");
    c.arity = 2; // arr, index
    c.local_count = 2;
    let arr = 0u16;
    let index = 1;

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op(Op::I32_CONST_1, 0); // deleteCount = 1
    let splice = imports.add_import("ecma:array", "splice");
    c.emit_op_u16(Op::CALL_IMPORT, splice, 0);
    c.emit(3u8, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op(Op::NULL, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_remove_value(arr, value) → bool ───────────────────────────────
// indexOf(arr, value) → if >= 0: splice(arr, idx, 1); return true; else return false.
fn build_array_remove_value(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_remove_value");
    c.arity = 2; // arr, value
    c.local_count = 3;
    let arr = 0u16;
    let value = 1;
    let idx = 2;

    // idx = indexOf(arr, value)
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    let index_of = imports.add_import("ecma:array", "indexOf");
    c.emit_op_u16(Op::CALL_IMPORT, index_of, 0);
    c.emit(2u8, 0);
    c.emit_op_u16(Op::LOCAL_SET, idx, 0);
    c.emit_op(Op::DROP, 0);

    // if idx >= 0: splice + return true
    c.emit_op_u16(Op::LOCAL_GET, idx, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GE, 0);
    let skip = c.emit_jump(Op::BR_IF_FALSE, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx, 0);
    c.emit_op(Op::I32_CONST_1, 0); // deleteCount = 1
    let splice = imports.add_import("ecma:array", "splice");
    c.emit_op_u16(Op::CALL_IMPORT, splice, 0);
    c.emit(3u8, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op(Op::TRUE, 0);
    c.emit_op(Op::RETURN, 0);

    c.patch_jump(skip);
    c.emit_op(Op::FALSE, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_insert_range(arr, index, src) → null ──────────────────────────
// Loop: for i in 0..src.length: splice(arr, index+i, 0, src[i])
fn build_array_insert_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_insert_range");
    c.arity = 3; // arr, index, src
    c.local_count = 5;
    let arr = 0u16; let index = 1; let src = 2; let i = 3; let src_len = 4;

    let len_import = imports.add_import("ecma:array", "length");
    let get_import = imports.add_import("ecma:array", "get");
    let splice_import = imports.add_import("ecma:array", "splice");

    // src_len = length(src)
    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    c.emit_op_u16(Op::CALL_IMPORT, len_import, 0); c.emit(1u8, 0);
    c.emit_op_u16(Op::LOCAL_SET, src_len, 0); c.emit_op(Op::DROP, 0);
    // i = 0
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    let blk = c.emit_block(0);
    let (lp, _) = c.emit_loop_s(0);
    // if i >= src_len break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, src_len, 0);
    c.emit_op(Op::DYN_GE, 0); c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);
    // splice(arr, index+i, 0, src[i])
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::CALL_IMPORT, get_import, 0); c.emit(2u8, 0);
    c.emit_op_u16(Op::CALL_IMPORT, splice_import, 0); c.emit(4u8, 0);
    c.emit_op(Op::DROP, 0);
    // i++
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(lp);
    c.emit_end(0); c.patch_block(blk);
    c.emit_op(Op::NULL, 0); c.emit_op(Op::RETURN, 0);
    c
}

// ── array_set_range(arr, index, src) → null ─────────────────────────────
// Loop: for i in 0..src.length: arr[index+i] = src[i]
fn build_array_set_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_set_range");
    c.arity = 3;
    c.local_count = 5;
    let arr = 0u16; let index = 1; let src = 2; let i = 3; let src_len = 4;

    let len_import = imports.add_import("ecma:array", "length");
    let get_import = imports.add_import("ecma:array", "get");
    let set_import = imports.add_import("ecma:array", "set");

    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    c.emit_op_u16(Op::CALL_IMPORT, len_import, 0); c.emit(1u8, 0);
    c.emit_op_u16(Op::LOCAL_SET, src_len, 0); c.emit_op(Op::DROP, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);

    let blk = c.emit_block(0);
    let (lp, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, src_len, 0);
    c.emit_op(Op::DYN_GE, 0); c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);
    // set(arr, index+i, get(src, i))
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::CALL_IMPORT, get_import, 0); c.emit(2u8, 0);
    c.emit_op_u16(Op::CALL_IMPORT, set_import, 0); c.emit(3u8, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0); c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(lp);
    c.emit_end(0); c.patch_block(blk);
    c.emit_op(Op::NULL, 0); c.emit_op(Op::RETURN, 0);
    c
}

// ── array_binary_search(arr, value) → i32 ───────────────────────────────
// Delegates to indexOf — correct for unsorted arrays, O(n) not O(log n)
// but avoids needing integer division opcode.
fn build_array_binary_search(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_binary_search");
    c.arity = 2; // arr, value
    c.local_count = 2;
    let arr = 0u16; let value = 1;
    let index_of = imports.add_import("ecma:array", "indexOf");
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, index_of, 0); c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_reverse_range(arr, index, count) → null ───────────────────────
fn build_array_reverse_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_reverse_range");
    c.arity = 3;
    c.local_count = 6;
    let arr = 0u16; let index = 1; let count = 2;
    let lo = 3; let hi = 4; let tmp = 5;

    let get_import = imports.add_import("ecma:array", "get");
    let set_import = imports.add_import("ecma:array", "set");

    // lo = index; hi = index + count - 1
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_SET, lo, 0); c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op(Op::I32_CONST_1, 0); c.emit_op(Op::DYN_NEG, 0); c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, hi, 0); c.emit_op(Op::DROP, 0);

    let blk = c.emit_block(0);
    let (lp, _) = c.emit_loop_s(0);
    // while lo < hi
    c.emit_op_u16(Op::LOCAL_GET, lo, 0);
    c.emit_op_u16(Op::LOCAL_GET, hi, 0);
    c.emit_op(Op::DYN_LT, 0); c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);
    // tmp = arr[lo]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, lo, 0);
    c.emit_op_u16(Op::CALL_IMPORT, get_import, 0); c.emit(2u8, 0);
    c.emit_op_u16(Op::LOCAL_SET, tmp, 0); c.emit_op(Op::DROP, 0);
    // arr[lo] = arr[hi]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, lo, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, hi, 0);
    c.emit_op_u16(Op::CALL_IMPORT, get_import, 0); c.emit(2u8, 0);
    c.emit_op_u16(Op::CALL_IMPORT, set_import, 0); c.emit(3u8, 0);
    c.emit_op(Op::DROP, 0);
    // arr[hi] = tmp
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, hi, 0);
    c.emit_op_u16(Op::LOCAL_GET, tmp, 0);
    c.emit_op_u16(Op::CALL_IMPORT, set_import, 0); c.emit(3u8, 0);
    c.emit_op(Op::DROP, 0);
    // lo++; hi--
    c.emit_op_u16(Op::LOCAL_GET, lo, 0);
    c.emit_op(Op::I32_CONST_1, 0); c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, lo, 0); c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, hi, 0);
    c.emit_op(Op::I32_CONST_1, 0); c.emit_op(Op::DYN_NEG, 0); c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, hi, 0); c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(lp);
    c.emit_end(0); c.patch_block(blk);
    c.emit_op(Op::NULL, 0); c.emit_op(Op::RETURN, 0);
    c
}

// ── array_last_index_of(arr, value) → i32 ───────────────────────────────
fn build_array_last_index_of(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_last_index_of");
    c.arity = 2; // arr, value
    c.local_count = 2;
    let arr = 0u16;
    let value = 1;

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    let last_index_of = imports.add_import("ecma:array", "lastIndexOf");
    c.emit_op_u16(Op::CALL_IMPORT, last_index_of, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── dir(path) → directory iterator ─────────────────────────
// PHP `dir($path)` returns a Directory object with `read()` / `close()`
// methods. Composes real WASI 0.2.8: get-directories →
// [method]descriptor.open-at (with OPEN_DIRECTORY=2) →
// [method]descriptor.read-directory → directory-entry-stream
// resource embedded on the wrapper as `__stream`. Each `read()` call
// pulls the next entry via [method]directory-entry-stream
// .read-directory-entry — proper lazy streaming, no upfront
// materialisation of the whole listing.
fn build_dir(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_dir");
    c.arity = 1; // path
    c.local_count = 3; // path(0) + descriptor(1) + stream(2)
    let path = 0u16;
    let descriptor = 1;
    let stream = 2;

    let get_directories = imports.add_import("wasi:filesystem/preopens", "get-directories");
    let open_at = imports.add_import("wasi:filesystem/types", "[method]descriptor.open-at");
    let read_directory = imports.add_import("wasi:filesystem/types", "[method]descriptor.read-directory");
    let array_get = imports.add_import("ecma:array", "get");
    let stream_key   = c.add_constant(Value::String(std::sync::Arc::from("__stream")));
    let read_key     = c.add_constant(Value::String(std::sync::Arc::from("read")));
    let close_key    = c.add_constant(Value::String(std::sync::Arc::from("close")));
    let read_global  = c.add_constant(Value::String(std::sync::Arc::from("__vybe_dir_read")));
    let close_global = c.add_constant(Value::String(std::sync::Arc::from("__vybe_dir_close")));
    let open_directory_flag = c.add_constant(Value::I32(2)); // open-flags::directory

    // preopen = get-directories()[0][0]
    c.emit_op_u16(Op::CALL_IMPORT, get_directories, 0); c.emit(0u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);

    // descriptor = preopen.open-at(0, path, OPEN_DIRECTORY, 0)
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, path, 0);
    c.emit_op_u16(Op::CONST, open_directory_flag, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, open_at, 0); c.emit(5u8, 0);
    c.emit_op_u16(Op::LOCAL_SET, descriptor, 0); c.emit_op(Op::DROP, 0);

    // stream = descriptor.read-directory()
    c.emit_op_u16(Op::LOCAL_GET, descriptor, 0);
    c.emit_op_u16(Op::CALL_IMPORT, read_directory, 0); c.emit(1u8, 0);
    c.emit_op_u16(Op::LOCAL_SET, stream, 0); c.emit_op(Op::DROP, 0);

    // obj = STRUCT_NEW 0
    c.emit_op_u16(Op::STRUCT_NEW, 0, 0);
    // obj.__stream = stream
    c.emit_op(Op::DUP, 0);
    c.emit_op_u16(Op::LOCAL_GET, stream, 0);
    c.emit_op_u16(Op::STRUCT_SET, stream_key, 0); c.emit_op(Op::DROP, 0);
    // obj.read = global __vybe_dir_read
    c.emit_op(Op::DUP, 0);
    c.emit_op_u16(Op::GLOBAL_GET, read_global, 0);
    c.emit_op_u16(Op::STRUCT_SET, read_key, 0); c.emit_op(Op::DROP, 0);
    // obj.close = global __vybe_dir_close
    c.emit_op(Op::DUP, 0);
    c.emit_op_u16(Op::GLOBAL_GET, close_global, 0);
    c.emit_op_u16(Op::STRUCT_SET, close_key, 0); c.emit_op(Op::DROP, 0);

    c.emit_op(Op::RETURN, 0);
    c
}

// ── dir.read(this) → next entry name or false ──────────────
// Pulls the next entry from the embedded WASI directory-entry-stream
// resource. End-of-stream (option<directory-entry>::none) maps to
// `false` for PHP `while ($f = $dir->read()) { … }` compatibility.
fn build_dir_read(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_dir_read");
    c.arity = 1; // this
    c.local_count = 1;
    let this = 0u16;

    let read_entry = imports.add_import(
        "wasi:filesystem/types",
        "[method]directory-entry-stream.read-directory-entry",
    );
    let stream_key = c.add_constant(Value::String(std::sync::Arc::from("__stream")));
    let name_key   = c.add_constant(Value::String(std::sync::Arc::from("name")));

    // entry = read-directory-entry(this.__stream)
    c.emit_op_u16(Op::LOCAL_GET, this, 0);
    c.emit_op_u16(Op::STRUCT_GET, stream_key, 0);
    c.emit_op_u16(Op::CALL_IMPORT, read_entry, 0); c.emit(1u8, 0);

    // if (entry === null) return false; else return entry.name
    c.emit_op(Op::DUP, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let if_null = c.emit_jump(Op::BR_IF_TRUE, 0);
    // not-null: pull `name` and return
    c.emit_op_u16(Op::STRUCT_GET, name_key, 0);
    c.emit_op(Op::RETURN, 0);
    // null: drop the null entry, return false
    c.patch_jump(if_null);
    c.emit_op(Op::DROP, 0);
    c.emit_op(Op::FALSE, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── dir.close(this) → null ─────────────────────────────────
// WASI's listDir returned the snapshot eagerly, so close is a no-op
// kept for source compatibility with `$dir->close();` calls.
fn build_dir_close(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_dir_close");
    c.arity = 1; // this
    c.local_count = 1;
    c.emit_op(Op::NULL, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── file(path) → array of lines ───────────────────────────
// PHP `file($path)` reads the entire file then splits on newline,
// returning the lines as an array (one element per line). Composes
// `wasi:filesystem.readFile` with `STR_SPLIT` — no host-side knowledge
// of the PHP-specific split semantic.
fn build_file(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_file");
    c.arity = 1; // path
    c.local_count = 1;

    let read_file = imports.add_import("wasi:filesystem", "readFile");
    let nl = c.add_constant(Value::String(std::sync::Arc::from("\n")));

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, read_file, 0); c.emit(1u8, 0);
    c.emit_op_u16(Op::CONST, nl, 0);
    c.emit_op(Op::STR_SPLIT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── file_exists(path) → bool ───────────────────────────────
// Calls stat-at; non-existent paths come back as an error object
// (no `type` field), every other result has a `type` string. So
// `result.type !== null` is a faithful exists-test.
fn build_file_exists(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_file_exists");
    c.arity = 1;
    c.local_count = 1;
    let get_directories = imports.add_import("wasi:filesystem/preopens", "get-directories");
    let stat_at = imports.add_import("wasi:filesystem/types", "[method]descriptor.stat-at");
    let array_get = imports.add_import("ecma:array", "get");
    let type_key = c.add_constant(Value::String(std::sync::Arc::from("type")));

    c.emit_op_u16(Op::CALL_IMPORT, get_directories, 0); c.emit(0u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, stat_at, 0); c.emit(3u8, 0);
    c.emit_op_u16(Op::STRUCT_GET, type_key, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── is_file(path) → bool ───────────────────────────────────
// stat-at then `type === "regular-file"`.
fn build_is_file(imports: &mut Chunk) -> Chunk {
    build_stat_type_match(imports, "__stdlib_is_file", "regular-file")
}

// ── is_dir(path) → bool ────────────────────────────────────
fn build_is_dir(imports: &mut Chunk) -> Chunk {
    build_stat_type_match(imports, "__stdlib_is_dir", "directory")
}

fn build_stat_type_match(imports: &mut Chunk, name: &str, expected_type: &str) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = 1;
    c.local_count = 1;
    let get_directories = imports.add_import("wasi:filesystem/preopens", "get-directories");
    let stat_at = imports.add_import("wasi:filesystem/types", "[method]descriptor.stat-at");
    let array_get = imports.add_import("ecma:array", "get");
    let type_key = c.add_constant(Value::String(std::sync::Arc::from("type")));
    let expected = c.add_constant(Value::String(std::sync::Arc::from(expected_type)));

    c.emit_op_u16(Op::CALL_IMPORT, get_directories, 0); c.emit(0u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, stat_at, 0); c.emit(3u8, 0);
    c.emit_op_u16(Op::STRUCT_GET, type_key, 0);
    c.emit_op_u16(Op::CONST, expected, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── filesize(path) → number ────────────────────────────────
// stat-at then read the `size` field. Errors return null (PHP
// returns false; close enough for now).
fn build_filesize(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_filesize");
    c.arity = 1;
    c.local_count = 1;
    let get_directories = imports.add_import("wasi:filesystem/preopens", "get-directories");
    let stat_at = imports.add_import("wasi:filesystem/types", "[method]descriptor.stat-at");
    let array_get = imports.add_import("ecma:array", "get");
    let size_key = c.add_constant(Value::String(std::sync::Arc::from("size")));

    c.emit_op_u16(Op::CALL_IMPORT, get_directories, 0); c.emit(0u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, stat_at, 0); c.emit(3u8, 0);
    c.emit_op_u16(Op::STRUCT_GET, size_key, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── unlink(path) → null ────────────────────────────────────
// PHP `unlink($path)` removes a file; under real WASI that's
// `[method]descriptor.unlink-file-at(parent, path)`. Return is
// null on success, error object on failure (matches WIT result<_>).
fn build_unlink(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_unlink");
    c.arity = 1;
    c.local_count = 1;
    let get_directories = imports.add_import("wasi:filesystem/preopens", "get-directories");
    let unlink_at = imports.add_import(
        "wasi:filesystem/types",
        "[method]descriptor.unlink-file-at",
    );
    let array_get = imports.add_import("ecma:array", "get");

    c.emit_op_u16(Op::CALL_IMPORT, get_directories, 0); c.emit(0u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, unlink_at, 0); c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── filemtime(path) → seconds since epoch ──────────────────
// PHP `filemtime($path)` returns Unix seconds. Composes real WASI
// 0.2.8 calls: get-directories → preopen descriptor → stat-at →
// data-modification-timestamp (ms) → divide by 1000.
fn build_filemtime(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_filemtime");
    c.arity = 1; // path
    c.local_count = 1;

    let get_directories = imports.add_import("wasi:filesystem/preopens", "get-directories");
    let stat_at = imports.add_import("wasi:filesystem/types", "[method]descriptor.stat-at");
    let array_get = imports.add_import("ecma:array", "get");
    let mtime_key = c.add_constant(Value::String(std::sync::Arc::from("data-modification-timestamp")));
    let one_thousand = c.add_constant(Value::F64(1000.0));

    // preopens[0][0] → descriptor for cwd
    c.emit_op_u16(Op::CALL_IMPORT, get_directories, 0); c.emit(0u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, array_get, 0); c.emit(2u8, 0);

    // stat-at(descriptor, path-flags=0, path) — flat path resolves
    // relative to the preopen, so PHP's "/abs/path" isn't supported
    // here; that's a Phase-2 issue (true preopen-aware path mapping).
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, stat_at, 0); c.emit(3u8, 0);

    c.emit_op_u16(Op::STRUCT_GET, mtime_key, 0);
    c.emit_op_u16(Op::CONST, one_thousand, 0);
    c.emit_op(Op::F64_DIV, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── Regex adapters for pattern-first language conventions ────────────
//
// PHP `preg_replace($pat, $repl, $str)` and Python `re.sub(pat, repl, str)`
// share the same `(pattern, replacement, input)` order. ECMA-262
// `String.prototype.replace` is `(input, regex, replacement)` (receiver
// first). The body just LOCAL_GETs in the right order then calls
// `ecma:regexp.replace`.

fn build_regex_replace_pat_first(imports: &mut Chunk) -> Chunk {
    let idx = imports.add_import("ecma:regexp", "replace");
    let mut c = Chunk::new("__stdlib_regex_replace_pat_first");
    c.arity = 3;
    c.local_count = 3; // pat(0), repl(1), str(2)
    // Push (str, pat, repl) — ecma:regexp.replace order
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(3u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// PHP `preg_split($pat, $str)` / Python `re.split(pat, str)` →
// `ecma:regexp.split(str, regex)`.
fn build_regex_split_pat_first(imports: &mut Chunk) -> Chunk {
    let idx = imports.add_import("ecma:regexp", "split");
    let mut c = Chunk::new("__stdlib_regex_split_pat_first");
    c.arity = 2;
    c.local_count = 2; // pat(0), str(1)
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// PHP `preg_match_all($pat, $str)` / Python `re.findall(pat, str)` →
// `ecma:regexp.matchAll(str, regex)`.
fn build_regex_match_all_pat_first(imports: &mut Chunk) -> Chunk {
    let idx = imports.add_import("ecma:regexp", "matchAll");
    let mut c = Chunk::new("__stdlib_regex_match_all_pat_first");
    c.arity = 2;
    c.local_count = 2; // pat(0), str(1)
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

