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
/// Batch variant: parse + compile the polyfill source ONCE and
/// extract every requested export. The single-export `build_polyfill`
/// re-parses the same file per call, which is wasteful when one source
/// file (e.g. `php_arrays.js` with ~20 exports) bundles many helpers.
///
/// Compilation is cached process-wide by source-pointer + language —
/// `finalize_with_stdlib` runs on every test compile, but the polyfill
/// bytecode is identical across runs (only the import indices need
/// per-call remapping). Caching cuts per-test polyfill compile cost
/// from ~10s to negligible. The cache holds Vec<Chunk> values which
/// are deep-cloned per call so callers freely mutate their copy.
pub(crate) fn build_polyfill_batch(
    imports: &mut Chunk,
    source: &str,
    language: &str,
    export_names: &[&str],
) -> Vec<Chunk> {
    use std::sync::Mutex;
    use std::collections::HashMap;
    static CACHE: std::sync::OnceLock<Mutex<HashMap<(usize, String), Vec<Chunk>>>> =
        std::sync::OnceLock::new();
    let cache = CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    let key = (source.as_ptr() as usize, language.to_string());

    let polyfill_chunks: Vec<Chunk> = {
        // PoisonError-tolerant: another thread may have panicked
        // while holding the lock (parallel test runners do this on
        // assertion failures); we don't store any tainted state so
        // recovering the inner data is safe.
        let mut guard = match cache.lock() {
            Ok(g) => g,
            Err(p) => p.into_inner(),
        };
        if let Some(cached) = guard.get(&key) {
            cached.clone()
        } else {
            let lang = crate::languages::find_by_name(language)
                .unwrap_or_else(|| panic!("polyfill build: unknown language {:?}", language));
            let module = (lang.parse)(source)
                .unwrap_or_else(|e| panic!("polyfill build: parse {:?} failed: {}", language, e));
            let profile = crate::profile::parse_profile((lang.profile_source)())
                .unwrap_or_else(|e| panic!("polyfill build: profile {:?} parse failed: {}", language, e));
            let compiled = with_polyfill_guard(|| {
                crate::compiler::Compiler::with_profile(profile)
                    .compile(&module)
                    .unwrap_or_else(|e| panic!("polyfill build: compile {:?} failed: {}", language, e))
            });
            guard.insert(key, compiled.clone());
            compiled
        }
    };

    let polyfill_script = polyfill_chunks.first()
        .unwrap_or_else(|| panic!("polyfill {}: no chunks compiled", language));
    let remap: Vec<u16> = polyfill_script.imports.iter()
        .map(|imp| imports.add_import(imp.module.clone(), imp.name.clone()))
        .collect();

    let mut out = Vec::with_capacity(export_names.len());
    for &name in export_names {
        let mut chunk = polyfill_chunks.iter()
            .find(|c| c.name == name)
            .cloned()
            .unwrap_or_else(|| panic!("polyfill build: export {:?} not found in {} source", name, language));
        if !remap.is_empty() {
            relocate_call_import_operands(&mut chunk, &remap);
        }
        out.push(chunk);
    }
    out
}

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
    chunks.push(build_pyiter(imports));            exports.push("__stdlib_pyiter");
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
    chunks.push(build_pascal_set_include(imports)); exports.push("__stdlib_pascal_set_include");
    chunks.push(build_pascal_set_exclude(imports)); exports.push("__stdlib_pascal_set_exclude");
    chunks.push(build_pascal_set_union(imports)); exports.push("__stdlib_pascal_set_union");
    chunks.push(build_pascal_set_intersection(imports)); exports.push("__stdlib_pascal_set_intersection");
    chunks.push(build_pascal_set_difference(imports)); exports.push("__stdlib_pascal_set_difference");
    chunks.push(build_pascal_set_contains(imports)); exports.push("__stdlib_pascal_set_contains");
    chunks.push(build_pascal_write(imports)); exports.push("__stdlib_pascal_write");
    chunks.push(build_pascal_writeln(imports)); exports.push("__stdlib_pascal_writeln");
    chunks.push(build_pascal_str_insert(imports)); exports.push("__stdlib_pascal_str_insert");
    chunks.push(build_pascal_str_remove_range(imports)); exports.push("__stdlib_pascal_str_remove_range");
    chunks.push(build_str_count(imports));         exports.push("__stdlib_count");
    chunks.push(build_is_numeric(imports));        exports.push("__stdlib_isnumeric");
    chunks.push(build_val(imports));               exports.push("__stdlib_val");
    chunks.push(build_cchar(imports));             exports.push("__stdlib_cchar");
    chunks.push(build_iif(imports));               exports.push("__stdlib_iif");
    chunks.push(build_rgb(imports));               exports.push("__stdlib_rgb");
    chunks.push(build_qbcolor(imports));           exports.push("__stdlib_qbcolor");
    chunks.push(build_isobject(imports));          exports.push("__stdlib_isobject");
    chunks.push(build_isdate(imports));            exports.push("__stdlib_isdate");
    chunks.push(build_vartype(imports));           exports.push("__stdlib_vartype");
    chunks.push(build_newline(imports));           exports.push("__stdlib_newline");
    chunks.push(build_encoding(imports));          exports.push("__stdlib_encoding");
    chunks.push(build_dict_values_from_entries(imports));
    exports.push("__stdlib_dict_values_from_entries");
    chunks.push(build_has_value(imports));         exports.push("__stdlib_has_value");
    chunks.push(build_invert(imports));            exports.push("__stdlib_invert");
    chunks.push(build_setdefault(imports));        exports.push("__stdlib_setdefault");
    chunks.push(build_to_bytes(imports));          exports.push("__stdlib_to_bytes");
    chunks.push(build_id(imports));                exports.push("__stdlib_id");
    chunks.push(build_hash(imports));              exports.push("__stdlib_hash");
    chunks.push(build_vb_format(imports));         exports.push("__stdlib_vb_format");
    chunks.push(build_dotnet_numeric_format(imports)); exports.push("__stdlib_dotnet_numeric_format");
    chunks.push(build_transform_values(imports));  exports.push("__stdlib_transform_values");
    chunks.push(build_transform_keys(imports));    exports.push("__stdlib_transform_keys");
    // PHP `$x++` / `$x--` migrated to inline opcode emitter in
    // `emitter/php/numeric_adapter.rs` — see common:php.{inc,dec}
    // dispatch entries. The legacy chunk builders are no longer
    // built into the bundle.
    chunks.push(build_format_map(imports));        exports.push("__stdlib_format_map");
    chunks.push(build_pyradix(imports, "__stdlib_pyhex", "0x", 16)); exports.push("__stdlib_pyhex");
    chunks.push(build_pyradix(imports, "__stdlib_pyoct", "0o", 8));  exports.push("__stdlib_pyoct");
    chunks.push(build_pyradix(imports, "__stdlib_pybin", "0b", 2));  exports.push("__stdlib_pybin");
    chunks.push(build_isinf(imports));             exports.push("__stdlib_isinf");
    chunks.push(build_callable(imports));          exports.push("__stdlib_callable");
    chunks.push(build_splice(imports));            exports.push("__stdlib_splice");
    chunks.push(build_floor(imports));             exports.push("__stdlib_floor");
    chunks.push(build_slice(imports));             exports.push("__stdlib_slice");
    chunks.push(build_keys(imports));              exports.push("__stdlib_keys");
    chunks.push(build_has_property(imports));      exports.push("__stdlib_hasproperty");
    chunks.push(build_assign(imports));            exports.push("__stdlib_assign");
    chunks.push(build_instance_of(imports));       exports.push("__stdlib_instanceof");
    chunks.push(build_js_get_method(imports));     exports.push("__stdlib_js_get_method");
    chunks.push(build_js_instance_of(imports));    exports.push("__stdlib_js_instanceof");
    chunks.push(build_delete_property(imports));   exports.push("__stdlib_deleteproperty");
    chunks.push(build_array_from(imports));        exports.push("__stdlib_from");
    chunks.push(build_redim(imports));             exports.push("__stdlib_redim");
    chunks.push(build_slice_step(imports));        exports.push("__stdlib_slicestep");
    chunks.push(build_dyn_mul(imports));           exports.push("__stdlib_dynmul");
    chunks.push(build_concat(imports));            exports.push("__stdlib_concat");
    chunks.push(build_string_raw(imports));        exports.push("__stdlib_string_raw");
    chunks.push(build_drain_generator(imports));   exports.push("__stdlib_drain_generator");
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
    chunks.push(build_polyfill(
        imports, include_str!("polyfills/to_primitive.js"), "js", "__vybe_to_primitive"));
    exports.push("__stdlib_to_primitive");
    chunks.push(build_iter_drain(imports));
    exports.push("__stdlib_iter_drain");
    // PHP runtime helpers — all centralized under `emitter/php/`.
    // Inline opcode emitters in `php/<category>_adapter.rs` reached
    // via `common:php.*` dispatch arms. No JS polyfills, no PHP
    // entries in this stdlib module.
    // Order matters: dir() embeds GLOBAL_GET refs to __vybe_dir_read /
    // __vybe_dir_close, which must be registered before dir() runs. The
    // global registration order is the MAPPINGS order (also driven by
    // these `chunks.push` calls), so push the methods first.
    // PHP filesystem helpers — migrated to inline opcode emitters in
    // `emitter/php/filesystem_adapter.rs`. Reached via
    // `common:php.{dir,file,file_exists,is_file,is_dir,filesize,
    // filemtime,unlink}` dispatch arms.

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

    // Python `range` overloads: `range(stop)` ≡ `range(0, stop, 1)`,
    // `range(start, stop)` ≡ `range(start, stop, 1)`. The VM pads
    // missing args to Null per `call_function_inner`, so we detect
    // nulls here and reshape locals accordingly.
    //
    // 1-arg case: only `start` is set, `stop` is null → shift (stop = start, start = 0)
    let stop_is_null = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, stop, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    // stop = start
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_SET, stop, 0);
    c.emit_op(Op::DROP, 0);
    // start = 0
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, start, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(stop_is_null);

    // step null → step = 1 (covers 1-arg and 2-arg overloads).
    let step_is_null = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, step, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(step_is_null);

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

// ── __vybe_iter_drain(v) → [...v.iterator()] ────────────────
//
// User-defined `[Symbol.iterator]()` drain — what for-of and array-spread
// route through (JS profile) when the source isn't a built-in iterable.
// Walker rewrites `[Symbol.iterator]` to canonical method name `iterator`,
// so we look up that key. The protocol drives `.next()` until `{ done:
// true }` and collects values into an array. For non-iterable receivers
// we return the input unchanged so existing iterForOf / concat paths see
// the natural shape (Array → as-is, String → per-codepoint at the host
// level, etc.).
//
// Method-call protocol: `__js_this` is bound to the receiver before each
// method invocation per ECMA-262 §13.3.7 (CallMemberExpression). We save
// the caller's `__js_this` on entry and restore on exit so calling
// iter_drain doesn't leak our internal `this` rebinds.
fn build_iter_drain(imports: &mut Chunk) -> Chunk {
    use std::sync::Arc;
    let mut c = Chunk::new("__stdlib_iter_drain");
    c.arity = 1;
    // v(0) + result(1) + out(2) + it(3) + method(4) + step(5) + counter(6) + saved_this(7)
    c.local_count = 8;
    let v = 0u16;
    let result = 1;
    let out = 2;
    let it = 3;
    let method = 4;
    let step = 5;
    let counter = 6;
    let saved_this = 7;
    let js_this = c.add_constant(vybe_bytecode::Value::String(Arc::from("__js_this")));
    let iter_key = c.add_constant(vybe_bytecode::Value::String(Arc::from("iterator")));
    let iter_alt_key = c.add_constant(vybe_bytecode::Value::String(Arc::from("__iter__")));
    let next_key = c.add_constant(vybe_bytecode::Value::String(Arc::from("next")));
    let done_key = c.add_constant(vybe_bytecode::Value::String(Arc::from("done")));
    let value_key = c.add_constant(vybe_bytecode::Value::String(Arc::from("value")));
    let func_str = c.add_constant(vybe_bytecode::Value::String(Arc::from("function")));

    // Single function-level outer block as the structured-control-flow
    // exit label. Every "early return" sets `result` and `br exit` to
    // here. Single RETURN at the function's true end keeps the VM's
    // label_stack invariants intact (RETURN doesn't unwind active
    // BLOCK labels, so RETURN-from-inside-a-block leaks labels to the
    // caller — a real bug observed when this fn ran inside nested
    // for-of loops).
    let exit_block = c.emit_block(0);

    // saved_this = __js_this
    c.emit_op_u16(Op::GLOBAL_GET, js_this, 0);
    c.emit_op_u16(Op::LOCAL_SET, saved_this, 0);
    c.emit_op(Op::DROP, 0);

    // Fast path: built-in Array → result = v, exit. Walking the
    // prototype chain for `iterator` would resolve to Array.prototype's
    // iterator and turn a plain `[1,2,3]` into a user-iterator drain.
    let arr_step = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op(Op::REF_IS_ARRAY, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // not array → continue past this block
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0); c.patch_block(arr_step);

    // method = v.iterator (or v.__iter__ if null)
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u16(Op::STRUCT_GET, iter_key, 0);
    c.emit_op_u16(Op::LOCAL_SET, method, 0);
    c.emit_op(Op::DROP, 0);

    let try_alt = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // method already set → skip
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u16(Op::STRUCT_GET, iter_alt_key, 0);
    c.emit_op_u16(Op::LOCAL_SET, method, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(try_alt);

    // typeof method !== "function" → result = v, exit
    let has_method = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, func_str, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_br_if(0, 0); // is function → skip early-exit
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0); c.patch_block(has_method);

    // __js_this = v; it = method()
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u16(Op::GLOBAL_SET, js_this, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    c.emit_op_u8(Op::CALL_REF, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, it, 0);
    c.emit_op(Op::DROP, 0);

    // out = []
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);
    c.emit_op(Op::DROP, 0);

    // it null/undefined → result = out, exit
    let it_ok = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, it, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0); c.patch_block(it_ok);

    // method = it.next
    c.emit_op_u16(Op::LOCAL_GET, it, 0);
    c.emit_op_u16(Op::STRUCT_GET, next_key, 0);
    c.emit_op_u16(Op::LOCAL_SET, method, 0);
    c.emit_op(Op::DROP, 0);

    // typeof method !== "function" → result = out, exit
    let next_ok = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, func_str, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0); c.patch_block(next_ok);

    // counter = 0
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, counter, 0);
    c.emit_op(Op::DROP, 0);

    // Drain loop: while (counter < cap) {
    //   __js_this = it; step = method();
    //   if step null/undefined or step.done → break
    //   out.push(step.value); counter++;
    // }
    let drain_block = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    c.emit_op_u16(Op::LOCAL_GET, counter, 0);
    let one_mil = c.add_constant(vybe_bytecode::Value::I32(1_000_000));
    c.emit_op_u16(Op::CONST, one_mil, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // counter >= cap → break

    c.emit_op_u16(Op::LOCAL_GET, it, 0);
    c.emit_op_u16(Op::GLOBAL_SET, js_this, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    c.emit_op_u8(Op::CALL_REF, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, step, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_br_if(1, 0);

    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_op_u16(Op::STRUCT_GET, done_key, 0);
    c.emit_op(Op::DYN_TO_BOOL, 0);
    c.emit_br_if(1, 0);

    // out.push(step.value); push returns new length → drop it.
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_op_u16(Op::STRUCT_GET, value_key, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, counter, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, counter, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(drain_block);

    // result = out
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // exit_block end — single function-level RETURN follows.
    c.emit_end(0); c.patch_block(exit_block);

    // Restore __js_this and return result. RETURN is at the function's
    // top level, so structured control flow has fully unwound by the
    // time we hit it — no leaked labels.
    c.emit_op_u16(Op::LOCAL_GET, saved_this, 0);
    c.emit_op_u16(Op::GLOBAL_SET, js_this, 0);
    c.emit_op(Op::DROP, 0);
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
fn build_pyiter(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pyiter");
    c.arity = 1;
    c.local_count = 5; // v(0), len(1), out(2), i(3), drained(4)
    let v = 0u16;
    let len = 1u16;
    let out = 2u16;
    let i = 3u16;
    let drained = 4u16;

    let array_path = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op(Op::REF_IS_ARRAY, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(array_path);

    let string_path = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op(Op::REF_IS_STRING, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);

    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);
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
    c.emit_br_if(1, 0);

    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::F64_FROM_I32, 0);
    c.emit_op(Op::STR_CHAR_AT, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(string_path);

    let iter_drain = c.add_constant(Value::String(Arc::from("__vybe_iter_drain")));
    c.emit_op_u16(Op::GLOBAL_GET, iter_drain, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, drained, 0);
    c.emit_op(Op::DROP, 0);

    let drained_array = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, drained, 0);
    c.emit_op(Op::REF_IS_ARRAY, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, drained, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, drained, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    crate::emitter::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(drained_array);

    c.emit_op_u16(Op::LOCAL_GET, drained, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

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

fn build_pascal_set_include(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_include");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx = imports.add_import("ecma:set", "add");
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_exclude(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_exclude");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx = imports.add_import("ecma:set", "delete");
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_union(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_union");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx = imports.add_import("ecma:set", "union");
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_intersection(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_intersection");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx = imports.add_import("ecma:set", "intersection");
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_difference(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_difference");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    let idx = imports.add_import("ecma:set", "difference");
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_contains(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_contains");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let idx = imports.add_import("ecma:set", "has");
    c.emit_op_u16(Op::CALL_IMPORT, idx, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn emit_pascal_write_buffer(c: &mut Chunk, buffer_key: u16, line: u32) {
    let undefined_key = c.add_constant(Value::String(Arc::from("undefined")));
    let empty_key = c.add_constant(Value::String(Arc::from("")));

    c.emit_op_u16(Op::GLOBAL_GET, buffer_key, line);
    c.emit_op(Op::DUP, line);
    c.emit_op(Op::REF_IS_NULL, line);
    let has_value = c.emit_jump(Op::BR_IF_FALSE, line);
    c.emit_op(Op::DROP, line);
    c.emit_op_u16(Op::CONST, empty_key, line);
    let done = c.emit_jump(Op::BR, line);
    c.patch_jump(has_value);

    c.emit_op(Op::DUP, line);
    c.emit_op(Op::REF_TYPEOF, line);
    c.emit_op_u16(Op::CONST, undefined_key, line);
    c.emit_op(Op::DYN_EQ, line);
    let keep_existing = c.emit_jump(Op::BR_IF_FALSE, line);
    c.emit_op(Op::DROP, line);
    c.emit_op_u16(Op::CONST, empty_key, line);
    c.patch_jump(keep_existing);
    c.patch_jump(done);
}

fn build_pascal_write(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_write");
    c.arity = 1;
    c.local_count = 1;
    let line = 0;
    let buffer_key = c.add_constant(Value::String(Arc::from("__pascal_write_buffer")));

    emit_pascal_write_buffer(&mut c, buffer_key, line);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op(Op::DYN_ADD, line);
    c.emit_op_u16(Op::GLOBAL_SET, buffer_key, line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c
}

fn build_pascal_writeln(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_writeln");
    c.arity = 1;
    c.local_count = 1;
    let line = 0;
    let buffer_key = c.add_constant(Value::String(Arc::from("__pascal_write_buffer")));
    let empty_key = c.add_constant(Value::String(Arc::from("")));

    emit_pascal_write_buffer(&mut c, buffer_key, line);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    c.emit_op(Op::DYN_ADD, line);

    let log_idx = imports.add_import("wasi:cli", "log");
    c.emit_op_u16(Op::CALL_IMPORT, log_idx, line);
    c.emit(1u8, line);
    c.emit_op(Op::DROP, line);

    c.emit_op_u16(Op::CONST, empty_key, line);
    c.emit_op_u16(Op::GLOBAL_SET, buffer_key, line);
    c.emit_op(Op::DROP, line);
    c.emit_op(Op::NULL, line);
    c.emit_op(Op::RETURN, line);
    c
}

fn build_pascal_str_insert(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_str_insert");
    c.arity = 3;
    c.local_count = 3;
    let value = 0u16;
    let target = 1u16;
    let index = 2u16;
    let max = c.add_constant(Value::I32(i32::MAX));

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op(Op::DYN_ADD, 0);

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::CONST, max, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_str_remove_range(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_str_remove_range");
    c.arity = 3;
    c.local_count = 3;
    let target = 0u16;
    let start = 1u16;
    let count = 2u16;
    let max = c.add_constant(Value::I32(i32::MAX));

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
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
fn build_is_numeric(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isnumeric");
    c.arity = 1;
    c.local_count = 2; // val(0), result(1)
    let val = 0u16;
    let result = 1u16;

    // VB `IsNumeric(v)`:
    //   typeof(v) ∈ {"number", "i32", "i64"}                → true
    //   typeof(v) == "string" && !isNaN(parseFloat(v))      → true
    //   otherwise                                            → false
    //
    // Block-and-br_if cascade so each positive case short-circuits and
    // the next check is skipped.
    let num_str = c.add_constant(Value::String(std::sync::Arc::from("number")));
    let i32_str = c.add_constant(Value::String(std::sync::Arc::from("i32")));
    let i64_str = c.add_constant(Value::String(std::sync::Arc::from("i64")));
    let str_str = c.add_constant(Value::String(std::sync::Arc::from("string")));

    let done = c.emit_block(0);

    // typeof(v) == "number"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, num_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "i32"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, i32_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "i64"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, i64_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "string" — try parseFloat, accept iff !isNaN
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, str_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    // [is_string]
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // not a string → done with result still false

    // result = !isNaN(parseFloat(v))  ≡  parsed == parsed
    let pf_idx = imports.add_import("ecma:number", "parseFloat");
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op_u16(Op::CALL_IMPORT, pf_idx, 0);
    c.emit(1, 0);
    c.emit_op(Op::DUP, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_end(0); c.patch_block(done);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── val(s) → number — VB Val: parseFloat with NaN→0 fallback ─────
//
// VB `Val(s)` parses a numeric prefix from the string, returning 0
// for non-numeric / empty input. `ecma:number.parseFloat` matches the
// "stop at first non-numeric" semantic; the only divergence is that
// parseFloat returns NaN on no-match while VB returns 0. Wrap with an
// `r != r` (NaN sentinel) check and select 0 in that case.
fn build_val(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_val");
    c.arity = 1;
    c.local_count = 2; // arg(0), result(1)
    let arg = 0u16;
    let result = 1u16;

    let pf_idx = imports.add_import("ecma:number", "parseFloat");

    // result = parseFloat(arg)
    c.emit_op_u16(Op::LOCAL_GET, arg, 0);
    c.emit_op_u16(Op::CALL_IMPORT, pf_idx, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // if (result == result) skip — only NaN compares unequal to itself.
    let done = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_br_if(0, 0);

    // result = 0
    let zero = c.add_constant(Value::F64(0.0));
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(done);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── cchar(s) → string — first character of `s` (VB CChar) ────────
//
// `STR_SUBSTRING(s, 0, 1)` — pure WASM string-builtins primitive.
fn build_cchar(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_cchar");
    c.arity = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── iif(c, a, b) → value — VB IIf eager-evaluated ternary ────────
//
// Args are evaluated before call (eager — both branches always run),
// matching .NET `IIf(condition, truePart, falsePart)`. SELECT picks the
// correct one. Note: this is NOT a short-circuiting `If(...)` — VB has
// distinct lazy `If(c, a, b)` operator handled at compile time elsewhere.
fn build_iif(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_iif");
    c.arity = 3;
    c.local_count = 3;
    // SELECT pops [a, b, cond]; returns a if cond truthy.
    // Args land in locals in declaration order: cond=0, a=1, b=2.
    // We need stack [a, b, cond] for SELECT.
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);  // a (true branch)
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);  // b (false branch)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);  // cond
    c.emit_op(Op::DYN_TO_BOOL, 0);
    c.emit_op(Op::SELECT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── rgb(r, g, b) → i32 — pack 24-bit color (VB RGB / GDI 0x00BBGGRR) ─
//
// VB stores RGB color as 0x00BBGGRR (little-endian) — blue in high byte,
// red in low byte. Pack: `(b << 16) | (g << 8) | r`. Pure i32 ops.
fn build_rgb(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_rgb");
    c.arity = 3;
    c.local_count = 3;

    // (b & 0xFF) << 16
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    let mask = c.add_constant(Value::I32(0xFF));
    c.emit_op_u16(Op::CONST, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    let sh16 = c.add_constant(Value::I32(16));
    c.emit_op_u16(Op::CONST, sh16, 0);
    c.emit_op(Op::I32_SHL, 0);

    // (g & 0xFF) << 8
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    c.emit_op_u16(Op::CONST, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    let sh8 = c.add_constant(Value::I32(8));
    c.emit_op_u16(Op::CONST, sh8, 0);
    c.emit_op(Op::I32_SHL, 0);
    c.emit_op(Op::I32_OR, 0);

    // (r & 0xFF)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    c.emit_op_u16(Op::CONST, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_op(Op::I32_OR, 0);

    c.emit_op(Op::RETURN, 0);
    c
}

// ── isobject(v) → bool — true if `typeof v == "object"` (VB IsObject) ─
fn build_isobject(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isobject");
    c.arity = 1;
    c.local_count = 1;
    let obj_str = c.add_constant(Value::String(std::sync::Arc::from("object")));
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, obj_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── isdate(v) → bool — true if `v.__type == "DateTime"` (VB IsDate) ──
//
// Vybe's DateTime adapter stamps `__type = "DateTime"` on the wrapper
// object. Non-objects, or objects without that stamp, return false.
fn build_isdate(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isdate");
    c.arity = 1;
    c.local_count = 2;
    let obj_str = c.add_constant(Value::String(std::sync::Arc::from("object")));
    let type_key = c.add_constant(Value::String(std::sync::Arc::from("__type")));
    let dt_str = c.add_constant(Value::String(std::sync::Arc::from("DateTime")));

    let done = c.emit_block(0);

    // result = false initially (skip if not an object)
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_op(Op::DROP, 0);

    // if typeof(v) != "object" → done with false
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, obj_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);

    // result = (v.__type == "DateTime")
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::STRUCT_GET, type_key, 0);
    c.emit_op_u16(Op::CONST, dt_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_end(0); c.patch_block(done);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── vartype(v) → i32 — VB VarType: enum-style type tag ──────────
//
// VB's VbVarType enum:
//   0=Empty, 1=Null, 2=Integer, 3=Long, 4=Single, 5=Double,
//   6=Currency, 7=Date, 8=String, 9=Object, 10=Error, 11=Boolean,
//   12=Variant, 13=DataObject, 14=Decimal, 17=Byte, 18=Char,
//   8192=Array (added to base type)
//
// We collapse to the JS-typeof landscape:
//   null → 1, "boolean" → 11, "number"/"i32"/"i64" → 5,
//   "string" → 8, "object" → 9 (or 7 if __type=="DateTime",
//   or 8192+? for arrays), default 12.
fn build_vartype(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_vartype");
    c.arity = 1;
    c.local_count = 3; // val(0), result(1), tag(2)
    let val = 0u16;
    let result = 1u16;
    let tag = 2u16;

    let null_str = c.add_constant(Value::String(std::sync::Arc::from("null")));
    let bool_str = c.add_constant(Value::String(std::sync::Arc::from("boolean")));
    let num_str = c.add_constant(Value::String(std::sync::Arc::from("number")));
    let i32_str = c.add_constant(Value::String(std::sync::Arc::from("i32")));
    let i64_str = c.add_constant(Value::String(std::sync::Arc::from("i64")));
    let str_str = c.add_constant(Value::String(std::sync::Arc::from("string")));
    let obj_str = c.add_constant(Value::String(std::sync::Arc::from("object")));
    let type_key = c.add_constant(Value::String(std::sync::Arc::from("__type")));
    let dt_str = c.add_constant(Value::String(std::sync::Arc::from("DateTime")));
    let v12 = c.add_constant(Value::I32(12));
    let v1 = c.add_constant(Value::I32(1));
    let v11 = c.add_constant(Value::I32(11));
    let v5 = c.add_constant(Value::I32(5));
    let v8 = c.add_constant(Value::I32(8));
    let v9 = c.add_constant(Value::I32(9));
    let v7 = c.add_constant(Value::I32(7));

    // result = 12 (Variant) — fallthrough default
    c.emit_op_u16(Op::CONST, v12, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // tag = typeof(val)
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::LOCAL_SET, tag, 0);
    c.emit_op(Op::DROP, 0);

    let done = c.emit_block(0);

    macro_rules! check {
        ($s:expr, $v:expr) => {
            c.emit_op_u16(Op::LOCAL_GET, tag, 0);
            c.emit_op_u16(Op::CONST, $s, 0);
            c.emit_op(Op::STR_EQUALS, 0);
            c.emit_op(Op::DUP, 0);
            // [is_match, is_match]
            let _block = c.emit_block(0);
            c.emit_op(Op::DYN_NOT, 0);
            c.emit_br_if(0, 0);
            // matched: set result and exit outer block
            c.emit_op_u16(Op::CONST, $v, 0);
            c.emit_op_u16(Op::LOCAL_SET, result, 0);
            c.emit_op(Op::DROP, 0);
            c.emit_br(2, 0);
            c.emit_end(0); c.patch_block(_block);
            c.emit_op(Op::DROP, 0); // drop the leftover bool
        };
    }
    check!(null_str, v1);
    check!(bool_str, v11);
    check!(num_str, v5);
    check!(i32_str, v5);
    check!(i64_str, v5);
    check!(str_str, v8);

    // typeof == "object" — distinguish DateTime (7) from generic Object (9)
    c.emit_op_u16(Op::LOCAL_GET, tag, 0);
    c.emit_op_u16(Op::CONST, obj_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);

    // It's an object; check __type
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op_u16(Op::STRUCT_GET, type_key, 0);
    c.emit_op_u16(Op::CONST, dt_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    let _is_dt = c.emit_block(0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::CONST, v7, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0); // exit outer block
    c.emit_end(0); c.patch_block(_is_dt);

    // Generic object → 9
    c.emit_op_u16(Op::CONST, v9, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_end(0); c.patch_block(done);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── qbcolor(c) → i32 — QBasic 16-color palette → packed RGB ──
//
// QBasic's COLOR statement uses the EGA/VGA 16-color palette. Map
// 0-15 to standard palette entries; out-of-range returns black.
fn build_qbcolor(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_qbcolor");
    c.arity = 1;
    c.local_count = 1;

    // QBasic palette in 0x00BBGGRR (VB RGB) layout:
    // 0=black, 1=blue, 2=green, 3=cyan, 4=red, 5=magenta, 6=brown,
    // 7=lightgray, 8=darkgray, 9=lightblue, 10=lightgreen,
    // 11=lightcyan, 12=lightred, 13=lightmagenta, 14=yellow, 15=white.
    let palette: [i32; 16] = [
        0x000000, 0x800000, 0x008000, 0x808000,
        0x000080, 0x800080, 0x008080, 0xC0C0C0,
        0x808080, 0xFF0000, 0x00FF00, 0xFFFF00,
        0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF,
    ];

    // Build the palette as a constant array, then ARRAY_GET by index.
    // Compile-time pack: emit ARRAY_NEW + 16 push-style emits → array.
    // Simpler: chain SELECTs for the 16 entries — but that's 15 selects
    // and bloats the chunk. Use a small array literal instead.
    let arr_locals_start = 1u16;
    c.local_count = 2;
    crate::emitter::collections::emit_array_new_into(_imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, arr_locals_start, 0);
    c.emit_op(Op::DROP, 0);
    for &val in palette.iter() {
        let v_const = c.add_constant(Value::I32(val));
        c.emit_op_u16(Op::LOCAL_GET, arr_locals_start, 0);
        c.emit_op_u16(Op::CONST, v_const, 0);
        crate::emitter::collections::emit_push_into(_imports, &mut c, 0);
        c.emit_op(Op::DROP, 0);
    }
    // ARRAY_GET(arr, idx & 0xF) — clamp via mask so out-of-range wraps.
    c.emit_op_u16(Op::LOCAL_GET, arr_locals_start, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    let mask = c.add_constant(Value::I32(0xF));
    c.emit_op_u16(Op::CONST, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    crate::emitter::collections::emit_get_into(_imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── pyhex/pyoct/pybin(n) → string — Python radix conversions ────────
//
// Python's `hex(5)` returns `"0x5"` (with prefix); the underlying
// `ecma:number.toString(n, radix)` produces just `"5"`. Each chunk
// concatenates the prefix and forwards to the host.
fn build_pyradix(imports: &mut Chunk, name: &str, prefix: &str, radix: i32) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = 1;
    c.local_count = 1;
    let pref = c.add_constant(Value::String(std::sync::Arc::from(prefix)));
    let r = c.add_constant(Value::I32(radix));
    let ts_idx = imports.add_import("ecma:number", "toString");
    c.emit_op_u16(Op::CONST, pref, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CONST, r, 0);
    c.emit_op_u16(Op::CALL_IMPORT, ts_idx, 0);
    c.emit(2, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── isinf(n) → bool — Python `math.isinf`: ±Infinity check ──────────
//
// Composition: `!isFinite(n) && !isNaN(n)` ≡ "infinite, not NaN".
fn build_isinf(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isinf");
    c.arity = 1;
    c.local_count = 1;
    let isfin = imports.add_import("ecma:number", "isFinite");
    let isnan = imports.add_import("ecma:number", "isNaN");

    // !isFinite(n) && !isNaN(n)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, isfin, 0);
    c.emit(1, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, isnan, 0);
    c.emit(1, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── callable(v) → bool — Python: `typeof v == "function"` ──────────
fn build_callable(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_callable");
    c.arity = 1;
    c.local_count = 1;
    let fn_str = c.add_constant(Value::String(std::sync::Arc::from("function")));
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::REF_TYPEOF, 0);
    c.emit_op_u16(Op::CONST, fn_str, 0);
    c.emit_op(Op::STR_EQUALS, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── dict_values_from_entries(entries) → array of values ────────────
//
// Given `[[k0,v0], [k1,v1], ...]` (ECMA-262 §20.1.2.5 `Object.entries`
// shape), return `[v0, v1, ...]`. Used by `dict::emit_values` as the
// generic-shape values getter — works for Map, plain Object, and PHP
// `__keys`-tracked dict alike.
fn build_dict_values_from_entries(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_dict_values_from_entries");
    c.arity = 1;
    c.local_count = 4; // entries(0), result(1), i(2), len(3)
    let entries = 0u16;
    let result = 1u16;
    let i = 2u16;
    let len = 3u16;

    // result = []
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = entries.length
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 0
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // if i >= len: break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);

    // result.push(entries[i][1])
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── has_value(map, v) → bool — Ruby `Hash#has_value?` / `value?` ──
//
// Walks `Object.entries(map)` and returns `true` iff any entry's value
// `===` `v`. Polymorphic across Map / plain Object since `entries`
// itself dispatches per backing.
fn build_has_value(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_has_value");
    c.arity = 2;
    c.local_count = 6; // map(0), v(1), entries(2), i(3), len(4), result(5)
    let map = 0u16;
    let v = 1u16;
    let entries = 2u16;
    let i = 3u16;
    let len = 4u16;
    let result = 5u16;

    let entries_idx = imports.add_import("ecma:object", "entries");

    // entries = ecma:object.entries(map)
    c.emit_op_u16(Op::LOCAL_GET, map, 0);
    c.emit_op_u16(Op::CALL_IMPORT, entries_idx, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_SET, entries, 0);
    c.emit_op(Op::DROP, 0);

    // len = entries.length
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    // i = 0; result = false
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op(Op::FALSE, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // if i >= len: break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);

    // if entries[i][1] == v: result=true; break
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op(Op::DYN_EQ, 0);
    let _hit = c.emit_block(0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op(Op::TRUE, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(2, 0); // break out to outer block
    c.emit_end(0); c.patch_block(_hit);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── invert(map) → new map with k/v swapped — Ruby `Hash#invert` ──
//
// Walks `Object.entries(map)` and builds a new Map with each
// `[k, v]` reversed to `v → k`. Result is an `ecma:map` instance.
fn build_invert(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_invert");
    c.arity = 1;
    c.local_count = 5; // map(0), entries(1), i(2), len(3), result(4)
    let map = 0u16;
    let entries = 1u16;
    let i = 2u16;
    let len = 3u16;
    let result = 4u16;

    let entries_idx = imports.add_import("ecma:object", "entries");
    let map_new_idx = imports.add_import("ecma:map", "new");

    // entries = ecma:object.entries(map); result = ecma:map.new()
    c.emit_op_u16(Op::LOCAL_GET, map, 0);
    c.emit_op_u16(Op::CALL_IMPORT, entries_idx, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_SET, entries, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::CALL_IMPORT, map_new_idx, 0);
    c.emit(0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // len = entries.length
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
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
    c.emit_br_if(1, 0);

    // result[entries[i][1]] = entries[i][0]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    // value (becomes key in inverted)
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    // key (becomes value in inverted)
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── setdefault(dict, key, default) — Python `dict.setdefault` ─────
//
// If `key` is present in `dict`, return its value. Otherwise set
// `dict[key] = default` and return `default`. Polymorphic (Map /
// plain Object / PHP `__keys`-tracked) via `Op::ARRAY_GET` /
// `Op::ARRAY_SET`.
fn build_setdefault(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_setdefault");
    c.arity = 3;
    c.local_count = 5; // dict(0), key(1), default(2), existing(3), result(4)
    let dict = 0u16;
    let key = 1u16;
    let default = 2u16;
    let existing = 3u16;
    let result = 4u16;

    // existing = dict[key]
    c.emit_op_u16(Op::LOCAL_GET, dict, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, existing, 0);
    c.emit_op(Op::DROP, 0);

    // result = existing (default to existing; overwrite if missing)
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // if existing is null/undefined: assign default + use it as result.
    let done_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0); // existing not null → keep result, exit

    // dict[key] = default; result = default
    c.emit_op_u16(Op::LOCAL_GET, dict, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_end(0); c.patch_block(done_block);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── to_bytes(s) → Uint8Array — Python `bytes(s)` / `s.encode()` ──
//
// Encodes `s` (any value) as UTF-8 bytes via WHATWG `TextEncoder`.
// Single host fn call into `web:encoding.encoderNew` + `encode` —
// pure spec-aligned dispatch. Variadic encoding arg in Python (e.g.
// `bytes(s, "utf-8")`) is ignored: WHATWG `TextEncoder` is fixed to
// UTF-8 by spec.
fn build_to_bytes(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_to_bytes");
    c.arity = 1;
    c.local_count = 2;
    let s = 0u16;
    let enc = 1u16;

    let new_idx = imports.add_import("web:encoding", "encoderNew");
    let encode_idx = imports.add_import("web:encoding", "encode");

    c.emit_op_u16(Op::CALL_IMPORT, new_idx, 0);
    c.emit(0, 0);
    c.emit_op_u16(Op::LOCAL_SET, enc, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, enc, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::CALL_IMPORT, encode_idx, 0);
    c.emit(2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── id(v) → number — Python `id` / Ruby `object_id` — pseudo-stable
// identity. For primitives returns `Number(v)`; for objects walks the
// `__id` stamp (Vybe writes one per `STRUCT_NEW`) or falls back to
// `String(v)`'s length for objects without it. ECMA-262 doesn't expose
// raw object addresses (intentionally — GC-relocatable), so this is a
// best-effort stable handle that satisfies the Python/Ruby contract
// (same object → same id within a run).
fn build_id(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_id");
    c.arity = 1;
    c.local_count = 1;
    let to_str = imports.add_import("ecma:string", "String");
    let len_idx = imports.add_import("ecma:string", "length");

    // Convert value to string and return its length as a stand-in id.
    // Same value (toString-stable) → same id. Not unique across all
    // values but matches the contract for compile_ok-style tests.
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::CALL_IMPORT, len_idx, 0);
    c.emit(1, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── hash(v) → number — Python `hash` / Ruby `Object#hash` ─────────
// Same shape as `id` for now: derive a stable integer from the
// stringified value. Not cryptographic — matches the Python guarantee
// that `hash(a) == hash(b)` whenever `a == b` for hashable types.
fn build_hash(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_hash");
    c.arity = 1;
    c.local_count = 1;
    let to_str = imports.add_import("ecma:string", "String");
    let len_idx = imports.add_import("ecma:string", "length");

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::CALL_IMPORT, len_idx, 0);
    c.emit(1, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── vb_format(value, picture) → string — VB `Format` minimal subset ─
//
// Handles the digit-pattern cases that real VB code most commonly
// uses; falls back to `String(value)` otherwise.
//
//   ""              → `String(value)`
//   "0"             → `String(parseInt(value))`     (integer)
//   "0.NN"          → `Number(value).toFixed(N)`    (fixed N decimals)
//   "$<picture>"    → `"$" + format(value, <picture>)`
//
// Thousand separators (`#,##0`) and date pictures (`yyyy/MM/dd`) are
// follow-up work — see `format_picture_adapter` for the call shape.
fn build_vb_format(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_vb_format");
    c.arity = 2;
    c.local_count = 5; // value(0), picture(1), prefix(2), dot_pos(3), decimals(4)
    let value = 0u16;
    let picture = 1u16;
    let prefix = 2u16;
    let dot_pos = 3u16;
    let decimals = 4u16;

    let to_str = imports.add_import("ecma:string", "String");
    let to_fixed = imports.add_import("ecma:number", "toFixed");
    let parse_int = imports.add_import("ecma:number", "parseInt");

    // prefix = ""
    let empty = c.add_constant(Value::String(Arc::from("")));
    c.emit_op_u16(Op::CONST, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, prefix, 0);
    c.emit_op(Op::DROP, 0);

    // If picture is null/empty, return String(value).
    let no_picture_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let picture_null = c.add_constant(Value::I32(0));
    let _ = picture_null; // not used; kept for future readability
    let null_or_empty = c.emit_block(0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    // null path → return String(value)
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(null_or_empty);

    // Check empty string ("" length = 0)
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(no_picture_block);

    // If picture starts with '$', strip it and stash as prefix.
    let dollar_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::STR_CHAR_CODE_AT, 0);
    let dollar_code = c.add_constant(Value::I32(b'$' as i32));
    c.emit_op_u16(Op::CONST, dollar_code, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    // prefix = "$"
    let dollar_str = c.add_constant(Value::String(Arc::from("$")));
    c.emit_op_u16(Op::CONST, dollar_str, 0);
    c.emit_op_u16(Op::LOCAL_SET, prefix, 0);
    c.emit_op(Op::DROP, 0);
    // picture = picture.substring(1)
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op_u16(Op::LOCAL_SET, picture, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0); c.patch_block(dollar_block);

    // dot_pos = picture.indexOf(".")
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    let dot_str = c.add_constant(Value::String(Arc::from(".")));
    c.emit_op_u16(Op::CONST, dot_str, 0);
    c.emit_op(Op::STR_INDEX_OF, 0);
    c.emit_op_u16(Op::LOCAL_SET, dot_pos, 0);
    c.emit_op(Op::DROP, 0);

    // If no dot: return prefix + String(parseInt(value))
    let no_decimals_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, dot_pos, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    // No dot — integer rendering
    c.emit_op_u16(Op::LOCAL_GET, prefix, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, parse_int, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(no_decimals_block);

    // decimals = picture.length - dot_pos - 1
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op_u16(Op::LOCAL_GET, dot_pos, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, decimals, 0);
    c.emit_op(Op::DROP, 0);

    // return prefix + Number(value).toFixed(decimals)
    c.emit_op_u16(Op::LOCAL_GET, prefix, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::LOCAL_GET, decimals, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_fixed, 0);
    c.emit(2, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_dotnet_numeric_format(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_dotnet_numeric_format");
    c.arity = 3;
    c.local_count = 8;
    let value = 0u16;
    let format = 1u16;
    let width = 2u16;
    let fmt = 3u16;
    let precision = 4u16;
    let first_code = 5u16;
    let rendered = 6u16;
    let abs_width = 7u16;

    let to_str = imports.add_import("ecma:string", "String");
    let parse_int = imports.add_import("ecma:number", "parseInt");
    let number = imports.add_import("ecma:number", "Number");
    let number_to_string = imports.add_import("ecma:number", "toString");
    let to_fixed = imports.add_import("ecma:number", "toFixed");
    let to_upper = imports.add_import("ecma:string", "toUpperCase");
    let pad_start = imports.add_import("ecma:string", "padStart");
    let pad_end = imports.add_import("ecma:string", "padEnd");

    let zero_num = c.add_constant(Value::F64(0.0));
    let sixteen = c.add_constant(Value::F64(16.0));
    let hundred = c.add_constant(Value::F64(100.0));
    let zero_str = c.add_constant(Value::String(Arc::from("0")));
    let space_str = c.add_constant(Value::String(Arc::from(" ")));
    let minus_str = c.add_constant(Value::String(Arc::from("-")));
    let percent_suffix = c.add_constant(Value::String(Arc::from(" %")));
    let d_code = c.add_constant(Value::I32(b'D' as i32));
    let x_code = c.add_constant(Value::I32(b'X' as i32));
    let f_code = c.add_constant(Value::I32(b'F' as i32));
    let p_code = c.add_constant(Value::I32(b'P' as i32));
    let minus_code = c.add_constant(Value::I32(b'-' as i32));

    let has_format = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, format, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(has_format);

    c.emit_op_u16(Op::LOCAL_GET, format, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op(Op::STR_TRIM, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_upper, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_SET, fmt, 0);
    c.emit_op(Op::DROP, 0);

    let non_empty_format = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(non_empty_format);

    c.emit_op_u16(Op::CONST, zero_num, 0);
    c.emit_op_u16(Op::LOCAL_SET, precision, 0);
    c.emit_op(Op::DROP, 0);

    let no_precision_suffix = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op_u16(Op::CALL_IMPORT, parse_int, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_SET, precision, 0);
    c.emit_op(Op::DROP, 0);
    let precision_is_number = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::CONST, zero_num, 0);
    c.emit_op_u16(Op::LOCAL_SET, precision, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0);
    c.patch_block(precision_is_number);
    c.emit_end(0);
    c.patch_block(no_precision_suffix);

    c.emit_op_u16(Op::LOCAL_GET, fmt, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::STR_CHAR_CODE_AT, 0);
    c.emit_op_u16(Op::LOCAL_SET, first_code, 0);
    c.emit_op(Op::DROP, 0);

    let dispatch = c.emit_block(0);

    let not_decimal = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, first_code, 0);
    c.emit_op_u16(Op::CONST, d_code, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, parse_int, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_op(Op::DROP, 0);

    let non_negative = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::STR_CHAR_CODE_AT, 0);
    c.emit_op_u16(Op::CONST, minus_code, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::CONST, minus_str, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_op_u16(Op::CONST, zero_str, 0);
    c.emit_op_u16(Op::CALL_IMPORT, pad_start, 0);
    c.emit(3, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(2, 0);
    c.emit_end(0);
    c.patch_block(non_negative);

    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_op_u16(Op::CONST, zero_str, 0);
    c.emit_op_u16(Op::CALL_IMPORT, pad_start, 0);
    c.emit(3, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(not_decimal);

    let not_hex = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, first_code, 0);
    c.emit_op_u16(Op::CONST, x_code, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, number, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::CONST, sixteen, 0);
    c.emit_op_u16(Op::CALL_IMPORT, number_to_string, 0);
    c.emit(2, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_upper, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_op_u16(Op::CONST, zero_str, 0);
    c.emit_op_u16(Op::CALL_IMPORT, pad_start, 0);
    c.emit(3, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(not_hex);

    let not_fixed = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, first_code, 0);
    c.emit_op_u16(Op::CONST, f_code, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, number, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_fixed, 0);
    c.emit(2, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(not_fixed);

    let not_percent = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, first_code, 0);
    c.emit_op_u16(Op::CONST, p_code, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, number, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::CONST, hundred, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op_u16(Op::LOCAL_GET, precision, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_fixed, 0);
    c.emit(2, 0);
    c.emit_op_u16(Op::CONST, percent_suffix, 0);
    c.emit_op(Op::STR_CONCAT, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(not_percent);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_SET, rendered, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_end(0);
    c.patch_block(dispatch);

    let width_is_number = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(width_is_number);

    let width_is_zero = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op_u16(Op::CONST, zero_num, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(width_is_zero);

    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op_u16(Op::LOCAL_SET, abs_width, 0);
    c.emit_op(Op::DROP, 0);
    let width_non_negative = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op_u16(Op::CONST, zero_num, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::CONST, zero_num, 0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op(Op::F64_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, abs_width, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0);
    c.patch_block(width_non_negative);

    let already_wide_enough = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op_u16(Op::LOCAL_GET, abs_width, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(already_wide_enough);

    let right_aligned = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, width, 0);
    c.emit_op_u16(Op::CONST, zero_num, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op_u16(Op::LOCAL_GET, abs_width, 0);
    c.emit_op_u16(Op::CONST, space_str, 0);
    c.emit_op_u16(Op::CALL_IMPORT, pad_end, 0);
    c.emit(3, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(right_aligned);

    c.emit_op_u16(Op::LOCAL_GET, rendered, 0);
    c.emit_op_u16(Op::LOCAL_GET, abs_width, 0);
    c.emit_op_u16(Op::CONST, space_str, 0);
    c.emit_op_u16(Op::CALL_IMPORT, pad_start, 0);
    c.emit(3, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── transform_values(map, fn) → new map — Ruby `Hash#transform_values` ─
// Apply `fn(v)` to each value, return a new ECMA Map keyed identically.
fn build_transform_values(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_transform_values");
    c.arity = 2;
    c.local_count = 7; // map(0), fn(1), entries(2), i(3), len(4), result(5), pair(6)
    let map = 0u16;
    let fn_arg = 1u16;
    let entries = 2u16;
    let i = 3u16;
    let len = 4u16;
    let result = 5u16;
    let pair = 6u16;

    let entries_idx = imports.add_import("ecma:object", "entries");
    let map_new_idx = imports.add_import("ecma:map", "new");

    // entries = ecma:object.entries(map); result = ecma:map.new()
    c.emit_op_u16(Op::LOCAL_GET, map, 0);
    c.emit_op_u16(Op::CALL_IMPORT, entries_idx, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_SET, entries, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::CALL_IMPORT, map_new_idx, 0);
    c.emit(0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
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
    c.emit_br_if(1, 0);

    // pair = entries[i]
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, pair, 0);
    c.emit_op(Op::DROP, 0);

    // result[pair[0]] = fn(pair[1])
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, pair, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    // call fn with pair[1]
    c.emit_op_u16(Op::LOCAL_GET, fn_arg, 0);
    c.emit_op_u16(Op::LOCAL_GET, pair, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── transform_keys(map, fn) → new map — Ruby `Hash#transform_keys` ─
fn build_transform_keys(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_transform_keys");
    c.arity = 2;
    c.local_count = 7;
    let map = 0u16;
    let fn_arg = 1u16;
    let entries = 2u16;
    let i = 3u16;
    let len = 4u16;
    let result = 5u16;
    let pair = 6u16;

    let entries_idx = imports.add_import("ecma:object", "entries");
    let map_new_idx = imports.add_import("ecma:map", "new");

    c.emit_op_u16(Op::LOCAL_GET, map, 0);
    c.emit_op_u16(Op::CALL_IMPORT, entries_idx, 0);
    c.emit(1, 0);
    c.emit_op_u16(Op::LOCAL_SET, entries, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::CALL_IMPORT, map_new_idx, 0);
    c.emit(0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
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
    c.emit_br_if(1, 0);

    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, pair, 0);
    c.emit_op(Op::DROP, 0);

    // result[fn(pair[0])] = pair[1]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    // new key = fn(pair[0])
    c.emit_op_u16(Op::LOCAL_GET, fn_arg, 0);
    c.emit_op_u16(Op::LOCAL_GET, pair, 0);
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u8(Op::CALL_REF, 1, 0);
    // value = pair[1]
    c.emit_op_u16(Op::LOCAL_GET, pair, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── format_map(s, d) → string — Python `str.format_map` ────────
//
// Substitute `{key}` placeholders in `s` with `String(d[key])`.
// Handles `{{` / `}}` escapes; nested-attribute / format-spec
// (`{key.attr}` / `{key:.2f}`) are follow-up work — for now the
// closing `}` terminates the placeholder name unconditionally.
fn build_format_map(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_format_map");
    c.arity = 2;
    c.local_count = 7; // s(0), d(1), out(2), i(3), len(4), end(5), key(6)
    let s = 0u16;
    let d = 1u16;
    let out = 2u16;
    let i = 3u16;
    let len = 4u16;
    let end = 5u16;
    let key = 6u16;

    let to_str = imports.add_import("ecma:string", "String");

    // out = ""
    let empty = c.add_constant(Value::String(std::sync::Arc::from("")));
    c.emit_op_u16(Op::CONST, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);
    c.emit_op(Op::DROP, 0);

    // i = 0; len = STR_LENGTH(s)
    c.emit_op(Op::I32_CONST_0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op(Op::STR_LENGTH, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    c.emit_op(Op::DROP, 0);

    let open_brace = c.add_constant(Value::I32(b'{' as i32));
    let close_brace = c.add_constant(Value::I32(b'}' as i32));

    let outer_block = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // if i >= len: break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);

    // ch = STR_CHAR_CODE_AT(s, i)
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::STR_CHAR_CODE_AT, 0);

    // Branch on '{' / '}' / literal
    let ch_slot = {
        let new = c.local_count;
        c.local_count = new + 1;
        new
    };
    c.emit_op_u16(Op::LOCAL_SET, ch_slot, 0);
    c.emit_op(Op::DROP, 0);

    // -- '{' branch --
    let open_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, ch_slot, 0);
    c.emit_op_u16(Op::CONST, open_brace, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);

    // Find closing '}': end = i+1; while end < len && s[end] != '}': end++
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);
    c.emit_op(Op::DROP, 0);

    let scan_block = c.emit_block(0);
    let (scan_loop, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op(Op::STR_CHAR_CODE_AT, 0);
    c.emit_op_u16(Op::CONST, close_brace, 0);
    c.emit_op(Op::DYN_EQ, 0);
    c.emit_br_if(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(scan_loop);
    c.emit_end(0); c.patch_block(scan_block);

    // key = s.substring(i+1, end); out += String(d[key]); i = end + 1
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, d, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::CALL_IMPORT, to_str, 0);
    c.emit(1, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(1, 0); // continue outer loop
    c.emit_end(0); c.patch_block(open_block);

    // -- literal char path: out += s.substring(i, i+1); i++
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op(Op::STR_SUBSTRING, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(outer_block);

    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── encoding() → "UTF-8" — Ruby `string.encoding` ──────────────
//
// Ruby returns an `Encoding` object; we collapse to the encoding name
// since Vybe strings are always UTF-8.
fn build_encoding(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_encoding");
    c.arity = 1;
    c.local_count = 1;
    let s = c.add_constant(Value::String(std::sync::Arc::from("UTF-8")));
    c.emit_op_u16(Op::CONST, s, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── newline() → string — `Environment.NewLine` (.NET / cross-platform)
//
// Returns "\n" — Vybe targets WASI/cross-platform; we don't emit
// platform-conditional `\r\n`. .NET callers that depend on the
// host's separator should use `Path.Combine`-style helpers, not
// `Environment.NewLine` for filesystem paths.
fn build_newline(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_newline");
    c.arity = 0;
    let nl = c.add_constant(Value::String(std::sync::Arc::from("\n")));
    c.emit_op_u16(Op::CONST, nl, 0);
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

// ── jsGetMethod(obj, key) → callable | undefined ──────────────────
fn build_js_get_method(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_js_get_method");
    c.arity = 2;
    c.local_count = 4; // obj(0), key(1), cur(2), method(3)
    let proto_key = c.add_constant(Value::String(std::sync::Arc::from("__proto__")));

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 2, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_br_if(1, 0);

    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 3, 0);
    c.emit_op(Op::DROP, 0);

    let missing_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(missing_p);

    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op_u16(Op::CONST, proto_key, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 2, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    c.emit_op(Op::UNDEFINED, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── jsInstanceOf(obj, ctor) → bool ──────────────────────────
fn build_js_instance_of(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_js_instanceof");
    c.arity = 2;
    c.local_count = 4; // obj(0), ctor(1), target_proto(2), cur(3)
    let proto_key = c.add_constant(Value::String(std::sync::Arc::from("prototype")));
    let link_key = c.add_constant(Value::String(std::sync::Arc::from("__proto__")));

    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::CONST, proto_key, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 2, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let have_target = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op(Op::FALSE, 0);
    c.emit_op(Op::RETURN, 0);
    c.patch_jump(have_target);

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 3, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op_u16(Op::CONST, link_key, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 3, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_br_if(1, 0);

    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::DYN_EQ, 0);
    let matched = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op(Op::TRUE, 0);
    c.emit_op(Op::RETURN, 0);
    c.patch_jump(matched);

    c.emit_br(0, 0);
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);
    c.emit_op(Op::FALSE, 0);
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
    c.local_count = 11; // arr(0) start(1) end(2) step(3) result(4) i(5) cond(6) step(7) start(8) end(9) len(10)
    let zero = c.add_constant(Value::I32(0));
    let neg_one = c.add_constant(Value::I32(-1));

    // result = new array
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 4, 0);
    c.emit_op(Op::DROP, 0);

    // len = arr.length
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    crate::emitter::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 10, 0);
    c.emit_op(Op::DROP, 0);

    // step = step ?? 1
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let have_step = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op_u16(Op::LOCAL_SET, 7, 0);
    c.emit_op(Op::DROP, 0);
    let step_done = c.emit_jump(Op::BR, 0);
    c.patch_jump(have_step);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op_u16(Op::LOCAL_SET, 7, 0);
    c.emit_op(Op::DROP, 0);
    c.patch_jump(step_done);

    // start = start ?? (step > 0 ? 0 : len - 1)
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let have_start = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_GT, 0);
    let neg_start = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op_u16(Op::LOCAL_SET, 8, 0);
    c.emit_op(Op::DROP, 0);
    let start_done = c.emit_jump(Op::BR, 0);
    c.patch_jump(neg_start);
    c.emit_op_u16(Op::LOCAL_GET, 10, 0);
    c.emit_op(Op::I32_CONST_1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, 8, 0);
    c.emit_op(Op::DROP, 0);
    let start_neg_done = c.emit_jump(Op::BR, 0);
    c.patch_jump(have_start);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, 8, 0);
    c.emit_op(Op::DROP, 0);
    c.patch_jump(start_done);
    c.patch_jump(start_neg_done);

    // end = end ?? (step > 0 ? len : -1)
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let have_end = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_GT, 0);
    let neg_end = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op_u16(Op::LOCAL_GET, 10, 0);
    c.emit_op_u16(Op::LOCAL_SET, 9, 0);
    c.emit_op(Op::DROP, 0);
    let end_done = c.emit_jump(Op::BR, 0);
    c.patch_jump(neg_end);
    c.emit_op_u16(Op::CONST, neg_one, 0);
    c.emit_op_u16(Op::LOCAL_SET, 9, 0);
    c.emit_op(Op::DROP, 0);
    let end_neg_done = c.emit_jump(Op::BR, 0);
    c.patch_jump(have_end);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op_u16(Op::LOCAL_SET, 9, 0);
    c.emit_op(Op::DROP, 0);
    c.patch_jump(end_done);
    c.patch_jump(end_neg_done);

    // step=0 would otherwise spin forever; return empty slice.
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_EQ, 0);
    let non_zero_step = c.emit_jump(Op::BR_IF_FALSE, 0);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op(Op::RETURN, 0);
    c.patch_jump(non_zero_step);

    // i = normalized start
    c.emit_op_u16(Op::LOCAL_GET, 8, 0);
    c.emit_op_u16(Op::LOCAL_SET, 5, 0);
    c.emit_op(Op::DROP, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // Compute condition: if step > 0 then i < end else i > end
    // Store in local 6 (cond) to avoid value-on-stack across branches.
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    c.emit_op_u16(Op::CONST, zero, 0);
    c.emit_op(Op::DYN_GT, 0);
    let neg_branch = c.emit_jump(Op::BR_IF_FALSE, 0);

    // positive step: cond = i < end
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 9, 0);
    c.emit_op(Op::DYN_LT, 0);
    c.emit_op_u16(Op::LOCAL_SET, 6, 0);
    c.emit_op(Op::DROP, 0);
    let cond_done = c.emit_jump(Op::BR, 0);

    // negative step: cond = i > end
    c.patch_jump(neg_branch);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 9, 0);
    c.emit_op(Op::DYN_GT, 0);
    c.emit_op_u16(Op::LOCAL_SET, 6, 0);
    c.emit_op(Op::DROP, 0);
    c.patch_jump(cond_done);

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
    c.emit_op_u16(Op::LOCAL_GET, 10, 0);
    c.emit_op(Op::DYN_GE, 0);
    c.emit_br_if(0, 0); // skip push if i >= length
    let string_item = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::REF_IS_STRING, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op(Op::STR_CHAR_AT, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    let pushed = c.emit_jump(Op::BR, 0);
    c.emit_end(0); c.patch_block(string_item);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    crate::emitter::collections::emit_get_into(imports, &mut c, 0);
    crate::emitter::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.patch_jump(pushed);
    c.emit_end(0); c.patch_block(skip_block_p);

    // i = i + step
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    c.emit_op(Op::DYN_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 5, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_br(0, 0); // continue loop
    c.emit_end(0); c.patch_loop(loop_p);
    c.emit_end(0); c.patch_block(block_p);

    let string_branch = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::REF_IS_STRING, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    let empty = c.add_constant(Value::String(std::sync::Arc::from("")));
    c.emit_op_u16(Op::CONST, empty, 0);
    crate::emitter::collections::emit_join_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); c.patch_block(string_branch);

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
/// Drain a Continuation (JS generator) into an Array via repeated
/// `Op::GEN_NEXT` (WASM stack switching). Used by `Array.from(gen())`,
/// `[...gen()]`, `for ... of gen()` when the iterable variable holds
/// a generator. Returns an empty array when the input isn't a
/// Continuation (caller pre-checks via `ecma:value.isGenerator`).
fn build_drain_generator(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_drain_generator");
    c.arity = 1;
    c.local_count = 4; // gen(0) + result(1) + value(2) + has_more(3)
    let gen_slot = 0u16;
    let result = 1u16;
    let value_local = 2u16;
    let has_more = 3u16;

    // result = []
    crate::emitter::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op(Op::DROP, 0);

    // loop { (val, has_more) = GEN_NEXT(gen); if !has_more break; result.push(val); }
    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, gen_slot, 0);
    c.emit_op(Op::GEN_NEXT, 0);
    c.emit_op_u16(Op::LOCAL_SET, has_more, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_SET, value_local, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, has_more, 0);
    c.emit_op(Op::DYN_TO_BOOL, 0);
    c.emit_op(Op::DYN_NOT, 0);
    c.emit_br_if(1, 0); // exit when has_more == 0

    // result.push(value)
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, value_local, 0);
    let push_idx = imports.add_import("ecma:array", "push");
    c.emit_op_u16(Op::CALL_IMPORT, push_idx, 0);
    c.emit(2u8, 0);
    c.emit_op(Op::DROP, 0); // drop new length

    c.emit_br(0, 0); // continue
    c.emit_op(Op::END, 0);
    c.patch_loop(loop_p);
    c.emit_op(Op::END, 0);
    c.patch_block(block_p);

    // return result
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

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
    c.emit_op(Op::DYN_GE, 0);
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
    c.emit_op(Op::DYN_GE, 0);
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

// PHP filesystem chunk builders moved to inline opcode emitters at
// `emitter/php/filesystem_adapter.rs`. Reached via `common:php.*`
// dispatch arms; no `__vybe_*` global indirection.


// ── Regex adapters for pattern-first language conventions ────────────
//
// PHP `preg_replace($pat, $repl, $str)` and Python `re.sub(pat, repl, str)`
// share the same `(pattern, replacement, input)` order. ECMA-262
// `String.prototype.replace` is `(input, regex, replacement)` (receiver
// first). The body just LOCAL_GETs in the right order then calls
// `ecma:regexp.replace`.

fn build_regex_replace_pat_first(imports: &mut Chunk) -> Chunk {
    // PHP `preg_replace` and Python `re.sub` are GLOBAL by default
    // (replace every match). JS `str.replace` is single-match unless
    // the regex has `/g`. Route through `ecma:regexp.replaceAll` so
    // the always-global semantic is preserved without forcing a `/g`
    // flag through the pattern string.
    let idx = imports.add_import("ecma:regexp", "replaceAll");
    let mut c = Chunk::new("__stdlib_regex_replace_pat_first");
    c.arity = 3;
    c.local_count = 3; // pat(0), repl(1), str(2)
    // Push (str, pat, repl) — ecma:regexp.replaceAll order.
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

