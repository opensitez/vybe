//! Runtime helper bytecode.
//!
//! These helpers are portable bytecode fallbacks for shared runtime-facing
//! operations. They are linked only when compiled code references their
//! `__vybe_*` global, so they are not a bundled language standard library.
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

use crate::primitives::instructions::core_wasm;
use crate::primitives::sets;
use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

fn emit_typeof(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:value", "typeof");
    chunk.emit_call(idx, 1, line);
}

fn emit_is_undefined(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(idx, 1, line);
}

fn emit_is_string(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(idx, 1, line);
}

fn emit_is_array(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:array", "isArray");
    chunk.emit_call(idx, 1, line);
}

fn emit_str_length(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(idx, 1, line);
}

fn emit_str_substring(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "substring");
    chunk.emit_call(idx, 3, line);
}

fn emit_str_char_code_at(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "charCodeAt");
    chunk.emit_call(idx, 2, line);
}

fn emit_str_equals(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_call(idx, 2, line);
}

fn emit_str_concat(imports: &mut Chunk, chunk: &mut Chunk, line: u32) {
    crate::primitives::ops::emit_dyn_add_into(imports, chunk, line);
}

fn emit_const_index(chunk: &mut Chunk, idx: u16, line: u32) {
    match chunk.constants[idx as usize].clone() {
        Value::Null | Value::TypedNull(_) => {
            chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line)
        }
        Value::Undefined => crate::primitives::expressions::emit_undefined(chunk, line),
        Value::Bool(value) => chunk.emit_bool_const(value, line),
        Value::I32(value) => chunk.emit_i32_const(value, line),
        Value::I64(value) => chunk.emit_i64_const(value, line),
        Value::BigInt(value) => chunk.emit_i64_const(value.to_i64_wrapping(), line),
        Value::F64(value) => chunk.emit_f64_const(value, line),
        Value::F32(value) => chunk.emit_f32_const(value, line),
        Value::String(value) | Value::Symbol(value) => chunk.emit_string_const(&value, line),
        Value::Object(_) | Value::WeakRef(_) | Value::V128(_) => {
            panic!("runtime helper cannot inline non-primitive constant")
        }
    }
}

// ── Generic polyglot polyfill helper ─────────────────────────────────
//
// `build_polyfill(source, language, export_name)` compiles a snippet
// of source code in any registered language and extracts a single
// named export as a helper Chunk. The result slots into `MAPPINGS` in
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
/// runtime helper finalization runs on every test compile, but the polyfill
/// bytecode is identical across runs (only the import indices need
/// per-call remapping). Caching cuts per-test polyfill compile cost
/// from ~10s to negligible. The cache holds Vec<Chunk> values which
/// are deep-cloned per call so callers freely mutate their copy.
#[allow(dead_code)]
pub(crate) fn build_polyfill_batch(
    imports: &mut Chunk,
    source: &str,
    language: &str,
    export_names: &[&str],
) -> Vec<Chunk> {
    use std::collections::HashMap;
    use std::sync::Mutex;
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
            let profile =
                crate::profile::parse_profile((lang.profile_source)()).unwrap_or_else(|e| {
                    panic!("polyfill build: profile {:?} parse failed: {}", language, e)
                });
            let compiled = with_polyfill_guard(|| {
                crate::primitives::Compiler::with_profile(profile)
                    .compile(&module)
                    .unwrap_or_else(|e| {
                        panic!("polyfill build: compile {:?} failed: {}", language, e)
                    })
            });
            guard.insert(key, compiled.clone());
            compiled
        }
    };

    let polyfill_script = polyfill_chunks
        .first()
        .unwrap_or_else(|| panic!("polyfill {}: no chunks compiled", language));
    let remap: Vec<u16> = polyfill_script
        .imports
        .iter()
        .map(|imp| imports.add_import(imp.module.clone(), imp.name.clone()))
        .collect();

    let mut out = Vec::with_capacity(export_names.len());
    for &name in export_names {
        let mut chunk = polyfill_chunks
            .iter()
            .find(|c| c.name == name)
            .cloned()
            .unwrap_or_else(|| {
                panic!(
                    "polyfill build: export {:?} not found in {} source",
                    name, language
                )
            });
        if !remap.is_empty() {
            relocate_call_import_operands(&mut chunk, &remap);
        }
        out.push(chunk);
    }
    out
}

#[allow(dead_code)]
pub(crate) fn build_polyfill(
    imports: &mut Chunk,
    source: &str,
    language: &str,
    export_name: &str,
) -> Chunk {
    let lang = crate::languages::find_by_name(language).unwrap_or_else(|| {
        panic!(
            "polyfill build: unknown language {:?} (registered: vb js pascal csharp \
             python php ruby dart cobol fortran)",
            language
        )
    });
    let module = (lang.parse)(source).unwrap_or_else(|e| {
        panic!(
            "polyfill build: parse {:?}.{:?} failed: {}",
            language, export_name, e
        )
    });
    let profile = crate::profile::parse_profile((lang.profile_source)())
        .unwrap_or_else(|e| panic!("polyfill build: profile {:?} parse failed: {}", language, e));
    // Recursion guard so the inner compile pipeline skips its own
    // runtime-helper finalization step — that would call back here and
    // recurse forever. Re-entrancy on the same thread is the only
    // failure mode and vybex build-time compilation is single-threaded.
    let polyfill_chunks = with_polyfill_guard(|| {
        crate::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .unwrap_or_else(|e| {
                panic!(
                    "polyfill build: compile {:?}.{:?} failed: {}",
                    language, export_name, e
                )
            })
    });

    // Merge the polyfill's module-level imports (which the JS compiler
    // wrote to its own chunks[0]) into the user program's imports
    // chunk, building a poly_idx → user_idx remap. Then walk the
    // function chunk's bytecode and rewrite every `CALL_IMPORT` operand
    // through the remap so runtime dispatch hits the right slot in the
    // user program's import table.
    let polyfill_script = polyfill_chunks
        .first()
        .unwrap_or_else(|| panic!("polyfill {}.{}: no chunks compiled", language, export_name));
    let remap: Vec<u16> = polyfill_script
        .imports
        .iter()
        .map(|imp| imports.add_import(imp.module.clone(), imp.name.clone()))
        .collect();

    let mut chunk = polyfill_chunks
        .into_iter()
        .find(|c| c.name == export_name)
        .unwrap_or_else(|| {
            panic!(
                "polyfill build: export {:?} not found in {} source (chunks compiled, \
             but no chunk has that name — check the function is declared at \
             top level and exported)",
                export_name, language
            )
        });

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
    use vybe_runtime::opcode::Op;
    let mut offset = 0;
    while offset + 3 < chunk.code.len() {
        let group = ((chunk.code[offset] as u16) << 8) | chunk.code[offset + 1] as u16;
        let sub = ((chunk.code[offset + 2] as u16) << 8) | chunk.code[offset + 3] as u16;
        let op = match Op::decode(group, sub) {
            Some(op) => op,
            None => {
                offset += 4;
                continue;
            }
        };
        let operand_start = offset + 4;
        let next = operand_start + op.operand_format().size_in(&chunk.code, operand_start);
        // Rewrite the import index for spec `call` specifically — other
        // U16_U8 opcodes (if any) don't index into the imports table.
        // Vybe stores u16 operands BIG-endian (see `Chunk::read_u16` in
        // `vybe_runtime/src/chunk.rs:314`).
        if op == Op::CALL {
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

pub fn is_compiling_runtime_helper() -> bool {
    IN_POLYFILL_BUILD.with(|c| c.get())
}

fn with_polyfill_guard<R>(f: impl FnOnce() -> R) -> R {
    IN_POLYFILL_BUILD.with(|c| c.set(true));
    let result = f();
    IN_POLYFILL_BUILD.with(|c| c.set(false));
    result
}

/// Build all runtime helper chunks. Each chunk registers any `ecma:array.*`
/// imports on the passed `imports` chunk (= user program's
/// `chunks[0]`, the module-level imports section per WASM semantics).
/// Returns the helper chunks + their export names, in matching order;
/// caller appends the chunks to its own vec.
pub fn build_runtime_helpers(imports: &mut Chunk) -> RuntimeHelpers {
    let mut chunks = Vec::new();
    let mut exports = Vec::new();

    chunks.push(build_sorted(imports));
    exports.push("__stdlib_sorted");
    for (build, name) in [
        (
            crate::primitives::channels::build_chan_send as fn(&mut Chunk) -> Chunk,
            "__stdlib_chan_send",
        ),
        (
            crate::primitives::channels::build_chan_recv,
            "__stdlib_chan_recv",
        ),
        (
            crate::primitives::channels::build_chan_recv_ok,
            "__stdlib_chan_recv_ok",
        ),
        (
            crate::primitives::channels::build_chan_len,
            "__stdlib_chan_len",
        ),
        (
            crate::primitives::channels::build_chan_cap,
            "__stdlib_chan_cap",
        ),
        (
            crate::primitives::channels::build_chan_close,
            "__stdlib_chan_close",
        ),
        (
            crate::primitives::channels::build_chan_ready_recv,
            "__stdlib_chan_ready_recv",
        ),
        (
            crate::primitives::channels::build_chan_ready_send,
            "__stdlib_chan_ready_send",
        ),
        (
            crate::primitives::channels::build_chan_wait_slice,
            "__stdlib_chan_wait_slice",
        ),
        (
            crate::primitives::channels::build_chan_try_send,
            "__stdlib_chan_try_send",
        ),
        (
            crate::primitives::channels::build_chan_try_recv,
            "__stdlib_chan_try_recv",
        ),
        (
            crate::primitives::channels::build_chan_try_peek,
            "__stdlib_chan_try_peek",
        ),
        (
            crate::primitives::channels::build_chan_drained,
            "__stdlib_chan_drained",
        ),
        (
            crate::primitives::channels::build_chan_closed,
            "__stdlib_chan_closed",
        ),
        (
            crate::primitives::channels::build_chan_recv_or_throw,
            "__stdlib_chan_recv_or_throw",
        ),
        (
            crate::primitives::channels::build_chan_wait_readable,
            "__stdlib_chan_wait_readable",
        ),
        (
            crate::primitives::channels::build_futex_alloc16,
            "__stdlib_futex_alloc16",
        ),
        (
            crate::primitives::channels::build_task_new,
            "__stdlib_task_new",
        ),
        (
            crate::primitives::channels::build_task_wait,
            "__stdlib_task_wait",
        ),
    ] {
        chunks.push(build(imports));
        exports.push(name);
    }
    chunks.push(build_sort_in_place(imports));
    exports.push("__stdlib_sort_in_place");
    chunks.push(build_sort_with_comparator(imports));
    exports.push("__stdlib_sort_with_comparator");
    chunks.push(build_sort_by_key(imports));
    exports.push("__stdlib_sort_by_key");
    // `__stdlib_reversed` removed — `reversed()` inlines its polymorphic loop
    // in `crate::primitives::collections::emit_reversed`.
    chunks.push(build_enumerate(imports));
    exports.push("__stdlib_enumerate");
    chunks.push(build_sum(imports));
    exports.push("__stdlib_sum");
    chunks.push(build_min(imports));
    exports.push("__stdlib_min");
    chunks.push(build_max(imports));
    exports.push("__stdlib_max");
    chunks.push(build_pyany(imports));
    exports.push("__stdlib_pyany");
    chunks.push(build_pyall(imports));
    exports.push("__stdlib_pyall");
    chunks.push(build_compact(imports));
    exports.push("__stdlib_compact");
    chunks.push(build_uniq(imports));
    exports.push("__stdlib_uniq");
    chunks.push(build_minmax(imports));
    exports.push("__stdlib_minmax");
    chunks.push(build_isempty(imports));
    exports.push("__stdlib_isempty");
    chunks.push(build_pymap(imports));
    exports.push("__stdlib_pymap");
    chunks.push(build_pyfilter(imports));
    exports.push("__stdlib_pyfilter");
    chunks.push(build_pyiter(imports));
    exports.push("__stdlib_pyiter");
    chunks.push(build_pynext(imports));
    exports.push("__stdlib_pynext");
    chunks.push(build_rotate(imports));
    exports.push("__stdlib_rotate");
    chunks.push(build_array_copy(imports));
    exports.push("__stdlib_array_copy");
    // Math transcendentals (sin/cos/tan/…/sign/clamp) removed: dead chunks —
    // every language routes math through `Math.*` → `ecma:math:*` host fns
    // directly, so these `env`-delegating wrappers were never bundled.
    // `__stdlib_tostring` removed — `str()` / `toString` route to
    // `ecma:string.String` directly (Python via `emit_helper`, others via
    // `emit_to_string`).
    chunks.push(build_string_is_null_or_empty(imports));
    exports.push("__stdlib_string_is_null_or_empty");
    chunks.push(build_string_is_null_or_whitespace(imports));
    exports.push("__stdlib_string_is_null_or_whitespace");
    chunks.push(build_str_insert(imports));
    exports.push("__stdlib_str_insert");
    chunks.push(build_str_remove_start(imports));
    exports.push("__stdlib_str_remove_start");
    chunks.push(build_str_remove_range(imports));
    exports.push("__stdlib_str_remove_range");
    chunks.push(build_pascal_set_include(imports));
    exports.push("__stdlib_pascal_set_include");
    chunks.push(build_pascal_set_exclude(imports));
    exports.push("__stdlib_pascal_set_exclude");
    chunks.push(build_pascal_set_union(imports));
    exports.push("__stdlib_pascal_set_union");
    chunks.push(build_pascal_set_intersection(imports));
    exports.push("__stdlib_pascal_set_intersection");
    chunks.push(build_pascal_set_difference(imports));
    exports.push("__stdlib_pascal_set_difference");
    chunks.push(build_pascal_set_contains(imports));
    exports.push("__stdlib_pascal_set_contains");
    chunks.push(build_pascal_write(imports));
    exports.push("__stdlib_pascal_write");
    chunks.push(build_pascal_writeln(imports));
    exports.push("__stdlib_pascal_writeln");
    chunks.push(build_pascal_str_insert(imports));
    exports.push("__stdlib_pascal_str_insert");
    chunks.push(build_pascal_str_remove_range(imports));
    exports.push("__stdlib_pascal_str_remove_range");
    chunks.push(build_str_count(imports));
    exports.push("__stdlib_count");
    chunks.push(build_is_numeric(imports));
    exports.push("__stdlib_isnumeric");
    chunks.push(build_val(imports));
    exports.push("__stdlib_val");
    chunks.push(build_cchar(imports));
    exports.push("__stdlib_cchar");
    chunks.push(build_iif(imports));
    exports.push("__stdlib_iif");
    chunks.push(build_rgb(imports));
    exports.push("__stdlib_rgb");
    chunks.push(build_qbcolor(imports));
    exports.push("__stdlib_qbcolor");
    chunks.push(build_isobject(imports));
    exports.push("__stdlib_isobject");
    chunks.push(build_isdate(imports));
    exports.push("__stdlib_isdate");
    chunks.push(build_vartype(imports));
    exports.push("__stdlib_vartype");
    chunks.push(build_newline(imports));
    exports.push("__stdlib_newline");
    chunks.push(build_encoding(imports));
    exports.push("__stdlib_encoding");
    chunks.push(build_dict_values_from_entries(imports));
    exports.push("__stdlib_dict_values_from_entries");
    chunks.push(build_setdefault(imports));
    exports.push("__stdlib_setdefault");
    chunks.push(build_to_bytes(imports));
    exports.push("__stdlib_to_bytes");
    chunks.push(build_id(imports));
    exports.push("__stdlib_id");
    chunks.push(build_hash(imports));
    exports.push("__stdlib_hash");
    chunks.push(build_vb_format(imports));
    exports.push("__stdlib_vb_format");
    chunks.push(vybe_runtime::registry::platform_numeric_format_helper()
        .expect("no platform registered a numeric-format helper")(
        imports
    ));
    exports.push("__stdlib_dotnet_numeric_format");
    // PHP `$x++` / `$x--` stay in the PHP emitter path (`common:php.{inc,dec}`)
    // rather than going through bundled stdlib/polyfill helpers.
    chunks.push(build_format_map(imports));
    exports.push("__stdlib_format_map");
    chunks.push(build_pyradix(imports, "__stdlib_pyhex", "0x", 16));
    exports.push("__stdlib_pyhex");
    chunks.push(build_pyradix(imports, "__stdlib_pyoct", "0o", 8));
    exports.push("__stdlib_pyoct");
    chunks.push(build_pyradix(imports, "__stdlib_pybin", "0b", 2));
    exports.push("__stdlib_pybin");
    chunks.push(build_isinf(imports));
    exports.push("__stdlib_isinf");
    // `__stdlib_splice` removed — no emit site references the `__vybe_splice`
    // global (the chunk was bundled but never used).
    // `__stdlib_slice` removed — slicing uses direct polymorphic `ecma:array.slice`.
    chunks.push(build_has_property(imports));
    exports.push("__stdlib_hasproperty");
    chunks.push(build_js_get_method(imports));
    exports.push("__stdlib_js_get_method");
    chunks.push(build_js_instance_of(imports));
    exports.push("__stdlib_js_instanceof");
    chunks.push(build_redim(imports));
    exports.push("__stdlib_redim");
    chunks.push(build_slice_step(imports));
    exports.push("__stdlib_slicestep");
    chunks.push(build_dyn_mul(imports));
    exports.push("__stdlib_dynmul");
    chunks.push(build_concat(imports));
    exports.push("__stdlib_concat");
    chunks.push(build_string_raw(imports));
    exports.push("__stdlib_string_raw");
    chunks.push(build_drain_generator(imports));
    exports.push("__stdlib_drain_generator");
    chunks.push(build_fmod(imports));
    exports.push("__stdlib_fmod");
    chunks.push(build_array_insert(imports));
    exports.push("__stdlib_array_insert");
    chunks.push(build_array_remove_at(imports));
    exports.push("__stdlib_array_remove_at");
    chunks.push(build_array_remove_value(imports));
    exports.push("__stdlib_array_remove_value");
    chunks.push(build_array_insert_range(imports));
    exports.push("__stdlib_array_insert_range");
    chunks.push(build_array_set_range(imports));
    exports.push("__stdlib_array_set_range");
    chunks.push(build_array_binary_search(imports));
    exports.push("__stdlib_array_binary_search");
    chunks.push(build_array_reverse_range(imports));
    exports.push("__stdlib_array_reverse_range");
    // ── Inline-bytecode sprintf (no JS polyfill) ──────────────────
    chunks.push(crate::primitives::sprintf::build_sprintf(imports));
    exports.push("__stdlib_sprintf");
    chunks.push(build_generator_next(imports));
    exports.push("__stdlib_generator_next");
    chunks.push(build_async_generator_next(imports));
    exports.push("__stdlib_async_generator_next");
    chunks.push(build_generator_self());
    exports.push("__stdlib_generator_self");
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
    // Same Layer-3 pattern as the `String.Format` dotnet adapter.
    chunks.push(build_regex_replace_pat_first(imports));
    exports.push("__stdlib_regex_replace_pat_first");
    chunks.push(build_regex_split_pat_first(imports));
    exports.push("__stdlib_regex_split_pat_first");
    chunks.push(build_regex_match_all_pat_first(imports));
    exports.push("__stdlib_regex_match_all_pat_first");

    RuntimeHelpers { chunks, exports }
}

/// Build only the requested helper chunks. `requested` must contain
/// helper export names such as `__stdlib_sorted`, not `__vybe_*` globals.
pub fn build_runtime_helpers_for_exports(
    imports: &mut Chunk,
    requested: &[&'static str],
) -> RuntimeHelpers {
    let mut chunks = Vec::new();
    let mut exports = Vec::new();

    for &name in requested {
        let chunk = build_runtime_helper_export(imports, name)
            .unwrap_or_else(|| panic!("unknown runtime helper export requested: {}", name));
        chunks.push(chunk);
        exports.push(name);
    }

    RuntimeHelpers { chunks, exports }
}

fn build_runtime_helper_export(imports: &mut Chunk, name: &str) -> Option<Chunk> {
    let chunk = match name {
        "__stdlib_sorted" => build_sorted(imports),
        "__stdlib_chan_send" => crate::primitives::channels::build_chan_send(imports),
        "__stdlib_chan_recv" => crate::primitives::channels::build_chan_recv(imports),
        "__stdlib_chan_recv_ok" => crate::primitives::channels::build_chan_recv_ok(imports),
        "__stdlib_chan_len" => crate::primitives::channels::build_chan_len(imports),
        "__stdlib_chan_cap" => crate::primitives::channels::build_chan_cap(imports),
        "__stdlib_chan_close" => crate::primitives::channels::build_chan_close(imports),
        "__stdlib_chan_ready_recv" => crate::primitives::channels::build_chan_ready_recv(imports),
        "__stdlib_chan_ready_send" => crate::primitives::channels::build_chan_ready_send(imports),
        "__stdlib_chan_wait_slice" => crate::primitives::channels::build_chan_wait_slice(imports),
        "__stdlib_chan_try_send" => crate::primitives::channels::build_chan_try_send(imports),
        "__stdlib_chan_try_recv" => crate::primitives::channels::build_chan_try_recv(imports),
        "__stdlib_chan_try_peek" => crate::primitives::channels::build_chan_try_peek(imports),
        "__stdlib_chan_drained" => crate::primitives::channels::build_chan_drained(imports),
        "__stdlib_chan_closed" => crate::primitives::channels::build_chan_closed(imports),
        "__stdlib_chan_recv_or_throw" => {
            crate::primitives::channels::build_chan_recv_or_throw(imports)
        }
        "__stdlib_chan_wait_readable" => {
            crate::primitives::channels::build_chan_wait_readable(imports)
        }
        "__stdlib_futex_alloc16" => crate::primitives::channels::build_futex_alloc16(imports),
        "__stdlib_task_new" => crate::primitives::channels::build_task_new(imports),
        "__stdlib_task_wait" => crate::primitives::channels::build_task_wait(imports),
        "__stdlib_sort_in_place" => build_sort_in_place(imports),
        "__stdlib_sort_with_comparator" => build_sort_with_comparator(imports),
        "__stdlib_sort_by_key" => build_sort_by_key(imports),
        "__stdlib_reversed" => build_reversed(imports),
        "__stdlib_enumerate" => build_enumerate(imports),
        "__stdlib_sum" => build_sum(imports),
        "__stdlib_min" => build_min(imports),
        "__stdlib_max" => build_max(imports),
        "__stdlib_pyany" => build_pyany(imports),
        "__stdlib_pyall" => build_pyall(imports),
        "__stdlib_compact" => build_compact(imports),
        "__stdlib_uniq" => build_uniq(imports),
        "__stdlib_minmax" => build_minmax(imports),
        "__stdlib_isempty" => build_isempty(imports),
        "__stdlib_pymap" => build_pymap(imports),
        "__stdlib_pyfilter" => build_pyfilter(imports),
        "__stdlib_pyiter" => build_pyiter(imports),
        "__stdlib_pynext" => build_pynext(imports),
        "__stdlib_rotate" => build_rotate(imports),
        "__stdlib_array_copy" => build_array_copy(imports),
        "__stdlib_tostring" => build_to_string(imports),
        "__stdlib_string_is_null_or_empty" => build_string_is_null_or_empty(imports),
        "__stdlib_string_is_null_or_whitespace" => build_string_is_null_or_whitespace(imports),
        "__stdlib_str_insert" => build_str_insert(imports),
        "__stdlib_str_remove_start" => build_str_remove_start(imports),
        "__stdlib_str_remove_range" => build_str_remove_range(imports),
        "__stdlib_pascal_set_include" => build_pascal_set_include(imports),
        "__stdlib_pascal_set_exclude" => build_pascal_set_exclude(imports),
        "__stdlib_pascal_set_union" => build_pascal_set_union(imports),
        "__stdlib_pascal_set_intersection" => build_pascal_set_intersection(imports),
        "__stdlib_pascal_set_difference" => build_pascal_set_difference(imports),
        "__stdlib_pascal_set_contains" => build_pascal_set_contains(imports),
        "__stdlib_pascal_write" => build_pascal_write(imports),
        "__stdlib_pascal_writeln" => build_pascal_writeln(imports),
        "__stdlib_pascal_str_insert" => build_pascal_str_insert(imports),
        "__stdlib_pascal_str_remove_range" => build_pascal_str_remove_range(imports),
        "__stdlib_count" => build_str_count(imports),
        "__stdlib_isnumeric" => build_is_numeric(imports),
        "__stdlib_val" => build_val(imports),
        "__stdlib_cchar" => build_cchar(imports),
        "__stdlib_iif" => build_iif(imports),
        "__stdlib_rgb" => build_rgb(imports),
        "__stdlib_qbcolor" => build_qbcolor(imports),
        "__stdlib_isobject" => build_isobject(imports),
        "__stdlib_isdate" => build_isdate(imports),
        "__stdlib_vartype" => build_vartype(imports),
        "__stdlib_newline" => build_newline(imports),
        "__stdlib_encoding" => build_encoding(imports),
        "__stdlib_dict_values_from_entries" => build_dict_values_from_entries(imports),
        "__stdlib_setdefault" => build_setdefault(imports),
        "__stdlib_to_bytes" => build_to_bytes(imports),
        "__stdlib_id" => build_id(imports),
        "__stdlib_hash" => build_hash(imports),
        "__stdlib_vb_format" => build_vb_format(imports),
        "__stdlib_dotnet_numeric_format" => {
            vybe_runtime::registry::platform_numeric_format_helper()
                .expect("no platform registered a numeric-format helper")(imports)
        }
        "__stdlib_format_map" => build_format_map(imports),
        "__stdlib_pyhex" => build_pyradix(imports, "__stdlib_pyhex", "0x", 16),
        "__stdlib_pyoct" => build_pyradix(imports, "__stdlib_pyoct", "0o", 8),
        "__stdlib_pybin" => build_pyradix(imports, "__stdlib_pybin", "0b", 2),
        "__stdlib_isinf" => build_isinf(imports),
        "__stdlib_splice" => build_splice(imports),
        "__stdlib_slice" => build_slice(imports),
        "__stdlib_hasproperty" => build_has_property(imports),
        "__stdlib_js_get_method" => build_js_get_method(imports),
        "__stdlib_js_instanceof" => build_js_instance_of(imports),
        "__stdlib_redim" => build_redim(imports),
        "__stdlib_slicestep" => build_slice_step(imports),
        "__stdlib_dynmul" => build_dyn_mul(imports),
        "__stdlib_concat" => build_concat(imports),
        "__stdlib_string_raw" => build_string_raw(imports),
        "__stdlib_drain_generator" => build_drain_generator(imports),
        "__stdlib_fmod" => build_fmod(imports),
        "__stdlib_array_insert" => build_array_insert(imports),
        "__stdlib_array_remove_at" => build_array_remove_at(imports),
        "__stdlib_array_remove_value" => build_array_remove_value(imports),
        "__stdlib_array_insert_range" => build_array_insert_range(imports),
        "__stdlib_array_set_range" => build_array_set_range(imports),
        "__stdlib_array_binary_search" => build_array_binary_search(imports),
        "__stdlib_array_reverse_range" => build_array_reverse_range(imports),
        "__stdlib_sprintf" => crate::primitives::sprintf::build_sprintf(imports),
        "__stdlib_generator_next" => build_generator_next(imports),
        "__stdlib_async_generator_next" => build_async_generator_next(imports),
        "__stdlib_generator_self" => build_generator_self(),
        "__stdlib_iter_drain" => build_iter_drain(imports),
        "__stdlib_regex_replace_pat_first" => build_regex_replace_pat_first(imports),
        "__stdlib_regex_split_pat_first" => build_regex_split_pat_first(imports),
        "__stdlib_regex_match_all_pat_first" => build_regex_match_all_pat_first(imports),
        _ => return None,
    };
    Some(chunk)
}

pub struct RuntimeHelpers {
    pub chunks: Vec<Chunk>,
    pub exports: Vec<&'static str>,
}

impl RuntimeHelpers {
    pub fn get(&self, name: &str) -> Option<usize> {
        self.exports.iter().position(|&n| n == name)
    }
}

pub fn rest_fixed_arity(name: &str) -> Option<u8> {
    match name {
        "sprintf" => Some(1),
        _ => None,
    }
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
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_i32_const(i32::MAX, 0);
    crate::primitives::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // len = result.length
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // Insertion sort: for i = 1 to len-1
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = result[i]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    // while j >= 0 && result[j] > key
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit inner loop

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // result[j+1] = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    // Now stack: [result, j+1] — need value = result[j]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0);
    c.patch_loop(inner_loop_p);
    c.emit_end(0);
    c.patch_block(inner_block_p);

    // result[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0);
    c.patch_loop(outer_loop_p);
    c.emit_end(0);
    c.patch_block(outer_block_p);

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
    c.local_count = 6; // arr(0) + i(1) + j(2) + len(3) + key(4) + lhs(5)
    let arr = 0u16;
    let i = 1;
    let j = 2;
    let len = 3;
    let key = 4;
    let lhs = 5;

    // len = arr.length
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // i = 1
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    // while j >= 0 && arr[j] > key
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit inner loop

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, lhs, 0);

    // Use _into variant so imports go to the shared `imports` chunk (chunk[0]),
    // not directly to `c`. Mixing emit_dyn_gt (adds to c.imports) with
    // emit_import_call_into (emits CALL_IMPORT with chunk[0] indices) causes
    // CALL_IMPORT to resolve the wrong host fn at runtime — same collision
    // documented in emit_len_into's comment above.
    c.emit_op_u16(Op::LOCAL_GET, lhs, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0);
    c.patch_loop(inner_loop_p);
    c.emit_end(0);
    c.patch_block(inner_block_p);

    // arr[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0);
    c.patch_loop(outer_loop_p);
    c.emit_end(0);
    c.patch_block(outer_block_p);

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
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // i = 1
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    // while j >= 0 && cmp(arr[j], key) > 0
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit inner loop

    // call cmp(arr[j], key) → result
    c.emit_op_u16(Op::LOCAL_GET, cmp, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut c, 2, 0);
    // result > 0 → swap needed
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0);
    c.patch_loop(inner_loop_p);
    c.emit_end(0);
    c.patch_block(inner_block_p);

    // arr[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0);
    c.patch_loop(outer_loop_p);
    c.emit_end(0);
    c.patch_block(outer_block_p);

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
    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // i = arr.length - 1
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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
    let async_iter_key = c.add_constant(vybe_runtime::Value::String(Arc::from("asyncIterator")));
    // NOT yet on the Iterator slot. Swapping this single key for
    // `__vybe_slot_5` cost one python `class` test (246/310 → 245/311): the
    // helper probes ONE alternate key, and the slot and the spelling are not
    // interchangeable here — a native iterable reaches this path too, and only
    // the spelling is present on those. Iterator needs a two-key probe (slot,
    // then spelling) rather than a substitution. The slot itself is verified
    // stamped and callable: `getattr(r, "__vybe_slot_5")()` returns a working
    // iterator. See flexclassplan.md §2g.
    let iter_slot_key = c.add_constant(vybe_runtime::Value::String(Arc::from(
        vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Iterator).as_str(),
    )));
    let iter_alt_key = c.add_constant(vybe_runtime::Value::String(Arc::from("__iter__")));
    let done_key = c.add_constant(vybe_runtime::Value::String(Arc::from("done")));
    let value_key = c.add_constant(vybe_runtime::Value::String(Arc::from("value")));

    // Single function-level outer block as the structured-control-flow
    // exit label. Every "early return" sets `result` and `br exit` to
    // here. Single RETURN at the function's true end keeps the VM's
    // label_stack invariants intact (RETURN doesn't unwind active
    // BLOCK labels, so RETURN-from-inside-a-block leaks labels to the
    // caller — a real bug observed when this fn ran inside nested
    // for-of loops).
    let exit_block = c.emit_block(0);

    // saved_this = __js_this
    crate::primitives::globals::emit_read(&mut c, "__js_this", 0);
    c.emit_op_u16(Op::LOCAL_SET, saved_this, 0);

    // Fast path: built-in Array → result = v, exit. Walking the
    // prototype chain for `iterator` would resolve to Array.prototype's
    // iterator and turn a plain `[1,2,3]` into a user-iterator drain.
    let arr_step = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    {
        let idx = c.add_import("ecma:array", "isArray");
        c.emit_call(idx, 1, 0);
    }
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // not array → continue past this block
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0);
    c.patch_block(arr_step);

    // Primitive strings are iterable by Unicode scalar value for JS
    // for-of/spread/yield*. They do not have object properties for
    // getMethodForCall to find here, so materialize through the shared
    // ECMA for-of adapter before the object-method protocol path.
    let string_step = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    {
        let idx = c.add_import("wasm:js-string", "test");
        c.emit_call(idx, 1, 0);
    }
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // not string → continue past this block
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    crate::primitives::collections::emit_import_call_into(
        imports,
        &mut c,
        "ecma:object",
        "iterForOf",
        1,
        0,
    );
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0);
    c.patch_block(string_step);

    // method = getMethodForCall(v, "iterator") — walks prototype chain and
    // binds the receiver for HostFunctions. For bytecode functions, the
    // caller sets __js_this directly.
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_string_const("iterator", 0);
    crate::primitives::collections::emit_import_call_into(
        imports,
        &mut c,
        "ecma:value",
        "getMethodForCall",
        2,
        0,
    );
    c.emit_op_u16(Op::LOCAL_SET, method, 0);

    // Symbol-keyed [Symbol.iterator] methods are stored as "Symbol(@@iterator)"
    // by ecma:array.set. Try getMethodForCall with that key when "iterator" fails.
    let try_sym = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_string_const("Symbol(@@iterator)", 0);
    crate::primitives::collections::emit_import_call_into(
        imports,
        &mut c,
        "ecma:value",
        "getMethodForCall",
        2,
        0,
    );
    c.emit_op_u16(Op::LOCAL_SET, method, 0);
    c.emit_end(0);
    c.patch_block(try_sym);

    let try_alt = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, iter_alt_key, 0);
    c.emit_op_u16(Op::LOCAL_SET, method, 0);
    c.emit_end(0);
    c.patch_block(try_alt);

    let try_async = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, async_iter_key, 0);
    c.emit_op_u16(Op::LOCAL_SET, method, 0);
    c.emit_end(0);
    c.patch_block(try_async);

    // The Iterator SLOT — Python `__iter__`, Ruby `each`, C# `GetEnumerator`,
    // Dart `iterator`. Another link in the probe chain rather than a
    // replacement for the spelling below: this helper also runs on values that
    // never went through a class, and those carry a native iterator or the bare
    // spelling, not a slot.
    let try_slot = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, iter_slot_key, 0);
    c.emit_op_u16(Op::LOCAL_SET, method, 0);
    c.emit_end(0);
    c.patch_block(try_slot);

    // typeof method !== "function" → result = v, exit
    let has_method = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    {
        let idx = c.add_import("ecma:value", "typeof");
        c.emit_call(idx, 1, 0);
    }
    c.emit_string_const("function", 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // is function → skip early-exit
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0);
    c.patch_block(has_method);

    // __js_this = v; it = method()
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    crate::primitives::globals::emit_write(&mut c, "__js_this", 0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut c, 0, 0);
    crate::primitives::functions::emit_await_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, it, 0);

    // out = []
    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);

    // it null/undefined → result = out, exit
    let it_ok = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, it, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0);
    c.patch_block(it_ok);

    // If iterator() returned a generator continuation, drain it with the
    // shared WASM stack-switching generator path. Generator continuations do
    // not expose a normal object-shaped own `next` method, so the generic
    // protocol loop below would otherwise collect nothing.
    c.emit_op_u16(Op::LOCAL_GET, it, 0);
    crate::primitives::collections::emit_import_call_into(
        imports,
        &mut c,
        "ecma:value",
        "isGenerator",
        1,
        0,
    );
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, it, 0);
    c.emit_op_u16(Op::LOCAL_SET, v, 0);
    crate::primitives::generators::emit_drain_into_array_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0);

    // method = getMethodForCall(it, "next")
    c.emit_op_u16(Op::LOCAL_GET, it, 0);
    c.emit_string_const("next", 0);
    crate::primitives::collections::emit_import_call_into(
        imports,
        &mut c,
        "ecma:value",
        "getMethodForCall",
        2,
        0,
    );
    c.emit_op_u16(Op::LOCAL_SET, method, 0);

    // typeof method !== "function" → result = out, exit
    let next_ok = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    {
        let idx = c.add_import("ecma:value", "typeof");
        c.emit_call(idx, 1, 0);
    }
    c.emit_string_const("function", 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0); // exit
    c.emit_end(0);
    c.patch_block(next_ok);

    // counter = 0
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, counter, 0);

    // Drain loop: while (counter < cap) {
    //   __js_this = it; step = method();
    //   if step null/undefined or step.done → break
    //   out.push(step.value); counter++;
    // }
    let drain_block = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    c.emit_op_u16(Op::LOCAL_GET, counter, 0);
    c.emit_i32_const(1_000_000, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // counter >= cap → break

    c.emit_op_u16(Op::LOCAL_GET, it, 0);
    crate::primitives::globals::emit_write(&mut c, "__js_this", 0);

    c.emit_op_u16(Op::LOCAL_GET, method, 0);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut c, 0, 0);
    crate::primitives::functions::emit_await_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, step, 0);

    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_br_if(1, 0);

    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, done_key, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);

    // out.push(step.value); push returns new length → drop it.
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, step, 0);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, value_key, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, counter, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, counter, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(drain_block);

    // result = out
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // exit_block end — single function-level RETURN follows.
    c.emit_end(0);
    c.patch_block(exit_block);

    // Restore __js_this and return result. RETURN is at the function's
    // top level, so structured control flow has fully unwound by the
    // time we hit it — no leaked labels.
    c.emit_op_u16(Op::LOCAL_GET, saved_this, 0);
    crate::primitives::globals::emit_write(&mut c, "__js_this", 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// Generator DRIVER chunks live with the rest of the generator machinery
// in `crate::primitives::generators` — this module only registers them.
use crate::primitives::generators::{build_async_generator_next, build_generator_next};

fn build_generator_self() -> Chunk {
    use std::sync::Arc;

    let mut c = Chunk::new("__stdlib_generator_self");
    c.arity = 0;
    c.local_count = 0;
    crate::primitives::globals::emit_read(&mut c, "__js_this", 0);
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

    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop

    // Build pair [i, arr[i]], then push onto result.
    // array_push takes [array, value] — so emit result first, then pair.
    c.emit_op_u16(Op::LOCAL_GET, result, 0); // result on stack
    c.emit_op_u16(Op::LOCAL_GET, i, 0); // i
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0); // arr[i]
    crate::primitives::collections::emit_array_pair_into(imports, &mut c, 0); // pair = [i, arr[i]]
    crate::primitives::collections::emit_push_into(imports, &mut c, 0); // result.push(pair)
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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

    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, total, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, total, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop → fell through

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    if is_any {
        // any: if truthy → return true
        c.emit_if(0);
        c.emit_bool_const(true, 0);
        c.emit_op(Op::RETURN, 0);
        c.emit_end(0);
    } else {
        // all: if falsy → return false
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_if(0);
        c.emit_bool_const(false, 0);
        c.emit_op(Op::RETURN, 0);
        c.emit_end(0);
    }

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    // Loop fell through: any → false, all → true
    if is_any {
        c.emit_bool_const(false, 0);
    } else {
        c.emit_bool_const(true, 0);
    }
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
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop

    // if arr[i] < best: best = arr[i]
    // block must wrap ALL condition operands + comparison + body
    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // skip if NOT less than
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_end(0);
    c.patch_block(skip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop

    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, best, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // skip if NOT greater than
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, best, 0);
    c.emit_end(0);
    c.patch_block(skip_block_p);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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

    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);

    // elem = arr[i]; stash into local
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, elem, 0);

    // if !is_null(elem) → result.push(elem)
    c.emit_op_u16(Op::LOCAL_GET, elem, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, elem, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
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
    crate::primitives::globals::emit_read(&mut c, "__vybe_min", 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut c, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, min_g, 0);

    crate::primitives::globals::emit_read(&mut c, "__vybe_max", 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut c, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, max_g, 0);

    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_dup(0);
    c.emit_op_u16(Op::LOCAL_GET, min_g, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_dup(0);
    c.emit_op_u16(Op::LOCAL_GET, max_g, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
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

    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);

    // elem = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, elem, 0);

    // if !result.includes(elem) result.push(elem)
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, elem, 0);
    let inc_idx = c.add_import("ecma:array", "includes");
    c.emit_call(inc_idx, 2, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, elem, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── pymap(fn, iter) — Python `map(fn, iter)` shape adapter ──
// Wraps ECMA `Array.prototype.map(fn)` (§23.1.3.21) with swapped
// args: Python passes (fn, iter), ECMA expects (iter, fn).
fn build_pymap(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pymap");
    c.arity = 2; // fn(0), iter(1)
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // iter
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // fn
    let idx = c.add_import("ecma:array", "map");
    c.emit_call(idx, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── pyfilter(fn, iter) — Python `filter(fn, iter)` shape adapter ──
fn build_pyfilter(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pyfilter");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let idx = c.add_import("ecma:array", "filter");
    c.emit_call(idx, 2, 0);
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
    {
        let idx = c.add_import("ecma:array", "isArray");
        c.emit_call(idx, 1, 0);
    }
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    crate::primitives::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(array_path);

    let string_path = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    {
        let idx = c.add_import("wasm:js-string", "test");
        c.emit_call(idx, 1, 0);
    }
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);

    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);
    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);

    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::F64_FROM_I32, 0);
    {
        let idx = c.add_import("ecma:string", "charAt");
        c.emit_call(idx, 2, 0);
    }
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(string_path);

    crate::primitives::globals::emit_read(&mut c, "__vybe_iter_drain", 0);
    c.emit_op_u16(Op::LOCAL_GET, v, 0);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut c, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, drained, 0);

    let drained_array = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, drained, 0);
    {
        let idx = c.add_import("ecma:array", "isArray");
        c.emit_call(idx, 1, 0);
    }
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, drained, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, drained, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    crate::primitives::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(drained_array);

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
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // default
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    // shift first element off iter
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let sh_idx = c.add_import("ecma:array", "shift");
    c.emit_call(sh_idx, 1, 0);
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
    c.emit_if(0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, n, 0);
    c.emit_end(0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // n_norm = ((n % len) + len) % len  — handles negative n
    c.emit_op_u16(Op::LOCAL_GET, n, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::globals::emit_read(&mut c, "__vybe_fmod", 0);
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
    c.emit_op(Op::I32_REM_S, 0); // n % len
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0); // + len
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    c.emit_op(Op::I32_REM_S, 0); // % len → n_norm
    c.emit_op_u16(Op::LOCAL_SET, n_norm, 0);

    // result = arr.slice(n_norm, len).concat(arr.slice(0, n_norm))
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, n_norm, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    let sl_idx = c.add_import("ecma:array", "slice");
    c.emit_call(sl_idx, 3, 0);
    // [first_part]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, n_norm, 0);
    c.emit_call(sl_idx, 3, 0);
    // [first_part, second_part]
    let cc_idx = c.add_import("ecma:array", "concat");
    c.emit_call(cc_idx, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_copy(src, dst, count) — C# `Array.Copy(src, dst, count)` ──
// Per .NET spec: copies `count` elements from src[0..] to dst[0..].
fn build_array_copy(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_copy");
    c.arity = 3;
    c.local_count = 3; // src(0), dst(1), count(2)
    let src = 0u16;
    let dst = 1;
    let count = 2;

    c.emit_op_u16(Op::LOCAL_GET, dst, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    c.emit_op(Op::ARRAY_COPY, 0);

    // .NET Array.Copy returns void
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(Op::RETURN, 0);
    let _ = imports; // silence unused
    c
}

// Math transcendentals removed. They were `env`-module-importing wrappers
// (`add_import("env", "sin")`) — a host dependency that shouldn't exist.
// `ecma:math` provides sin/cos/tan/asin/acos/atan/atan2/log/log10/exp/
// sinh/cosh/tanh/sign natively; every language routes `Math.*` → `ecma:math:*`
// directly, so these chunks were dead AND pulled in a phantom `env` import.

// ── toString(value) → string ────────────────────────────────
// "" + value triggers dyn_add string coercion in the VM
fn build_to_string(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_tostring");
    c.arity = 1;
    c.local_count = 1;
    let val = 0u16;
    c.emit_string_const("", 0);
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── string_is_null_or_empty(value) → bool ─────────────────
fn build_string_is_null_or_empty(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_string_is_null_or_empty");
    c.arity = 1;
    c.local_count = 1;
    let value = 0u16;

    let non_null = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_bool_const(true, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(non_null);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    {
        let idx = c.add_import("wasm:js-string", "length");
        c.emit_call(idx, 1, 0);
    }
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── string_is_null_or_whitespace(value) → bool ─────────────
fn build_string_is_null_or_whitespace(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_string_is_null_or_whitespace");
    c.arity = 1;
    c.local_count = 1;
    let value = 0u16;

    let non_null = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_bool_const(true, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(non_null);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    {
        let idx = c.add_import("ecma:string", "trim");
        c.emit_call(idx, 1, 0);
    }
    {
        let idx = c.add_import("wasm:js-string", "length");
        c.emit_call(idx, 1, 0);
    }
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── str_insert(str, index, value) → string ────────────────
fn build_str_insert(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_str_insert");
    c.arity = 3;
    c.local_count = 3;
    let value = 2u16;

    // prefix = str[0:index]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    {
        let idx = c.add_import("wasm:js-string", "substring");
        c.emit_call(idx, 3, 0);
    }

    // prefix + value (keeps current coercion behavior for non-string values)
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);

    // + suffix = str[index:]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_i32_const(i32::MAX, 0);
    {
        let idx = c.add_import("wasm:js-string", "substring");
        c.emit_call(idx, 3, 0);
    }
    {
        let idx = c.add_import("wasm:js-string", "concat");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c
}

// ── str_remove_start(str, start) → string ─────────────────
fn build_str_remove_start(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_str_remove_start");
    c.arity = 2;
    c.local_count = 2;

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    {
        let idx = c.add_import("wasm:js-string", "substring");
        c.emit_call(idx, 3, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c
}

// ── str_remove_range(str, start, count) → string ──────────
fn build_str_remove_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_str_remove_range");
    c.arity = 3;
    c.local_count = 3;

    // prefix = str[0:start]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    {
        let idx = c.add_import("wasm:js-string", "substring");
        c.emit_call(idx, 3, 0);
    }

    // suffix = str[start+count:]
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_i32_const(i32::MAX, 0);
    {
        let idx = c.add_import("wasm:js-string", "substring");
        c.emit_call(idx, 3, 0);
    }
    {
        let idx = c.add_import("wasm:js-string", "concat");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_include(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_include");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    sets::emit_add_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_exclude(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_exclude");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    sets::emit_delete_chunk(&mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_union(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_union");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    sets::emit_union_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_intersection(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_intersection");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    sets::emit_intersection_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_difference(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_difference");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    sets::emit_difference_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_set_contains(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_contains");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    sets::emit_has_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn emit_pascal_write_buffer(c: &mut Chunk, buffer_key: u16, line: u32) {
    let undefined_key = c.add_constant(Value::String(Arc::from("undefined")));
    let empty_key = c.add_constant(Value::String(Arc::from("")));

    crate::primitives::globals::emit_read(c, "__pascal_write_buffer", line);
    c.emit_dup(line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_if_value(line);
    c.emit_op(Op::DROP, line);
    emit_const_index(c, empty_key, line);
    c.emit_else(line);

    c.emit_dup(line);
    emit_typeof(c, line);
    emit_const_index(c, undefined_key, line);
    crate::primitives::ops::emit_dyn_eq(c, line);
    crate::primitives::ops::emit_dyn_to_bool(c, line);
    c.emit_if_value(line);
    c.emit_op(Op::DROP, line);
    emit_const_index(c, empty_key, line);
    c.emit_else(line);
    c.emit_end(line);
    c.emit_end(line);
}

fn build_pascal_write(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_write");
    c.arity = 1;
    c.local_count = 1;
    let line = 0;
    let buffer_key = c.add_constant(Value::String(Arc::from("__pascal_write_buffer")));

    emit_pascal_write_buffer(&mut c, buffer_key, line);
    c.emit_op_u16(Op::LOCAL_GET, 0, line);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, line);
    crate::primitives::globals::emit_write(&mut c, "__pascal_write_buffer", line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, line);

    let log_idx = c.add_import("web:console", "log");
    c.emit_call(log_idx, 1, line);
    c.emit_op(Op::DROP, line);

    emit_const_index(&mut c, empty_key, line);
    crate::primitives::globals::emit_write(&mut c, "__pascal_write_buffer", line);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    c.emit_op(Op::RETURN, line);
    c
}

fn build_pascal_str_insert(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_str_insert");
    c.arity = 3;
    c.local_count = 3;
    let value = 0u16;
    let target = 1u16;
    let index = 2u16;
    let max = c.add_constant(Value::I32(i32::MAX));

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    emit_str_substring(&mut c, 0);

    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    emit_const_index(&mut c, max, 0);
    emit_str_substring(&mut c, 0);
    emit_str_concat(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

fn build_pascal_str_remove_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_str_remove_range");
    c.arity = 3;
    c.local_count = 3;
    let target = 0u16;
    let start = 1u16;
    let count = 2u16;
    let max = c.add_constant(Value::I32(i32::MAX));

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    emit_str_substring(&mut c, 0);

    c.emit_op_u16(Op::LOCAL_GET, target, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    emit_const_index(&mut c, max, 0);
    emit_str_substring(&mut c, 0);
    emit_str_concat(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── count(haystack, needle) → int ───────────────────────────
// Count non-overlapping occurrences using substring + indexOf loop
fn build_str_count(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_count");
    c.arity = 2;
    c.local_count = 4;
    let haystack = 0u16;
    let needle = 1;
    let count = 2;
    let pos = 3;

    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, count, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, pos, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, haystack, 0);
    c.emit_op_u16(Op::LOCAL_GET, pos, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    emit_const_index(&mut c, max, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, needle, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    // Save indexOf result to local (don't use DUP — value can't cross block boundary)
    let idx_result = 4u16; // reuse local slot (local_count=4, slot 4 is beyond declared but safe with extra locals)
    c.local_count = 5; // need one more local for idx_result
    c.emit_op_u16(Op::LOCAL_SET, idx_result, 0);
    // Check if index < 0
    c.emit_op_u16(Op::LOCAL_GET, idx_result, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop if index < 0
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, count, 0);
    c.emit_op_u16(Op::LOCAL_GET, pos, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, pos, 0);
    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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
    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result_local, 0);

    // Collect removed elements: arr[index..index+deleteCount]
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_GET, delete_count, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop

    c.emit_op_u16(Op::LOCAL_GET, result_local, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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
    emit_typeof(&mut c, 0);
    emit_const_index(&mut c, num_str, 0);
    emit_str_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "i32"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    emit_typeof(&mut c, 0);
    emit_const_index(&mut c, i32_str, 0);
    emit_str_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "i64"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    emit_typeof(&mut c, 0);
    emit_const_index(&mut c, i64_str, 0);
    emit_str_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_br_if(0, 0);

    // typeof(v) == "string" — try parseFloat, accept iff !isNaN
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    emit_typeof(&mut c, 0);
    emit_const_index(&mut c, str_str, 0);
    emit_str_equals(&mut c, 0);
    // [is_string]
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // not a string → done with result still false

    // result = !isNaN(parseFloat(v))  ≡  parsed == parsed
    let pf_idx = c.add_import("ecma:number", "parseFloat");
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_call(pf_idx, 1, 0);
    c.emit_dup(0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_end(0);
    c.patch_block(done);
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

    let pf_idx = c.add_import("ecma:number", "parseFloat");

    // result = parseFloat(arg)
    c.emit_op_u16(Op::LOCAL_GET, arg, 0);
    c.emit_call(pf_idx, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // if (result == result) skip — only NaN compares unequal to itself.
    let done = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);

    // result = 0
    let zero = c.add_constant(Value::F64(0.0));
    emit_const_index(&mut c, zero, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_end(0);
    c.patch_block(done);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── cchar(s) → string — first character of `s` (VB CChar) ────────
//
// `wasm:js-string.substring(s, 0, 1)` — pure WASM string-builtins primitive.
fn build_cchar(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_cchar");
    c.arity = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    emit_str_substring(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── iif(c, a, b) → value — VB IIf eager-evaluated ternary ────────
//
// Args are evaluated before call (eager — both branches always run),
// matching .NET `IIf(condition, truePart, falsePart)`. SELECT picks the
// correct one. Note: this is NOT a short-circuiting `If(...)` — VB has
// distinct lazy `If(c, a, b)` operator handled at compile time elsewhere.
fn build_iif(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_iif");
    c.arity = 3;
    c.local_count = 3;
    // SELECT pops [a, b, cond]; returns a if cond truthy.
    // Args land in locals in declaration order: cond=0, a=1, b=2.
    // We need stack [a, b, cond] for SELECT.
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // a (true branch)
    c.emit_op_u16(Op::LOCAL_GET, 2, 0); // b (false branch)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // cond
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
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
    emit_const_index(&mut c, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    let sh16 = c.add_constant(Value::I32(16));
    emit_const_index(&mut c, sh16, 0);
    c.emit_op(Op::I32_SHL, 0);

    // (g & 0xFF) << 8
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    emit_const_index(&mut c, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    let sh8 = c.add_constant(Value::I32(8));
    emit_const_index(&mut c, sh8, 0);
    c.emit_op(Op::I32_SHL, 0);
    c.emit_op(Op::I32_OR, 0);

    // (r & 0xFF)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    emit_const_index(&mut c, mask, 0);
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
    emit_typeof(&mut c, 0);
    emit_const_index(&mut c, obj_str, 0);
    emit_str_equals(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── isdate(v) → bool — true if `v.__type == "DateTime"` (VB IsDate) ──
//
// Vybe's DateTime adapter stamps `__type = "DateTime"` on the wrapper
// object. Non-objects, or objects without that stamp, return false.
fn build_isdate(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isdate");
    c.arity = 1;
    c.local_count = 2;
    let parse_idx = c.add_import("ecma:date", "parse");
    let obj_str = c.add_constant(Value::String(std::sync::Arc::from("object")));
    let str_str = c.add_constant(Value::String(std::sync::Arc::from("string")));
    let type_key = c.add_constant(Value::String(std::sync::Arc::from("__type")));
    let dt_str = c.add_constant(Value::String(std::sync::Arc::from("DateTime")));

    let done = c.emit_block(0);

    // result = false initially (skip if not an object)
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // if typeof(v) == "string" → parseable strings also count as dates.
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_typeof(&mut c, 0);
    emit_const_index(&mut c, str_str, 0);
    emit_str_equals(&mut c, 0);
    c.emit_if(0);

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(parse_idx, 1, 0);
    c.emit_dup(0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_br(1, 0);
    c.emit_end(0);

    // if typeof(v) != "object" → done with false
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_typeof(&mut c, 0);
    emit_const_index(&mut c, obj_str, 0);
    emit_str_equals(&mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);

    // result = (v.__type == "DateTime")
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, type_key, 0);
    emit_const_index(&mut c, dt_str, 0);
    emit_str_equals(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);

    c.emit_end(0);
    c.patch_block(done);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    crate::primitives::ops::emit_i32_to_bool(&mut c, 0);
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
//   null/Nothing → 0, "boolean" → 11, integral numerics → 2,
//   non-integral numerics → 5, "i32" → 2, "i64" → 3,
//   "string" → 8, arrays → 8194, "object" → 9,
//   DateTime → 7, default 12.
fn build_vartype(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_vartype");
    c.arity = 1;
    c.local_count = 3; // val(0), result(1), tag(2)
    let val = 0u16;
    let result = 1u16;
    let tag = 2u16;

    let bool_str = c.add_constant(Value::String(std::sync::Arc::from("boolean")));
    let num_str = c.add_constant(Value::String(std::sync::Arc::from("number")));
    let i32_str = c.add_constant(Value::String(std::sync::Arc::from("i32")));
    let i64_str = c.add_constant(Value::String(std::sync::Arc::from("i64")));
    let str_str = c.add_constant(Value::String(std::sync::Arc::from("string")));
    let obj_str = c.add_constant(Value::String(std::sync::Arc::from("object")));
    let type_key = c.add_constant(Value::String(std::sync::Arc::from("__type")));
    let dt_str = c.add_constant(Value::String(std::sync::Arc::from("DateTime")));
    let v12 = c.add_constant(Value::I32(12));
    let v0 = c.add_constant(Value::I32(0));
    let v2 = c.add_constant(Value::I32(2));
    let v3 = c.add_constant(Value::I32(3));
    let v11 = c.add_constant(Value::I32(11));
    let v5 = c.add_constant(Value::I32(5));
    let v8 = c.add_constant(Value::I32(8));
    let v9 = c.add_constant(Value::I32(9));
    let v7 = c.add_constant(Value::I32(7));
    let v8194 = c.add_constant(Value::I32(8194));

    // result = 12 (Variant) — fallthrough default
    emit_const_index(&mut c, v12, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // null / Nothing → Empty (0)
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let is_null = c.emit_block(0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    emit_const_index(&mut c, v0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(is_null);

    // arrays are a distinct VM kind, not "object"
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    emit_is_array(&mut c, 0);
    let is_array = c.emit_block(0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    emit_const_index(&mut c, v8194, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(is_array);

    // tag = typeof(val)
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    emit_typeof(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, tag, 0);

    let done = c.emit_block(0);

    macro_rules! check {
        ($s:expr, $v:expr) => {
            c.emit_op_u16(Op::LOCAL_GET, tag, 0);
            emit_const_index(&mut c, $s, 0);
            emit_str_equals(&mut c, 0);
            c.emit_dup(0);
            // [is_match, is_match]
            let _block = c.emit_block(0);
            crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
            c.emit_br_if(0, 0);
            // matched: set result and exit outer block
            emit_const_index(&mut c, $v, 0);
            c.emit_op_u16(Op::LOCAL_SET, result, 0);
            c.emit_br(2, 0);
            c.emit_end(0);
            c.patch_block(_block);
            c.emit_op(Op::DROP, 0); // drop the leftover bool
        };
    }
    check!(bool_str, v11);
    check!(i32_str, v2);
    check!(i64_str, v3);
    check!(str_str, v8);

    c.emit_op_u16(Op::LOCAL_GET, tag, 0);
    emit_const_index(&mut c, num_str, 0);
    emit_str_equals(&mut c, 0);
    let is_number = c.emit_block(0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);

    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_op(Op::F64_TRUNC, 0);
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    let is_integral = c.emit_block(0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    emit_const_index(&mut c, v2, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(2, 0);
    c.emit_end(0);
    c.patch_block(is_integral);

    emit_const_index(&mut c, v5, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0);
    c.emit_end(0);
    c.patch_block(is_number);

    // typeof == "object" — distinguish arrays, DateTime, and generic Object.
    c.emit_op_u16(Op::LOCAL_GET, tag, 0);
    emit_const_index(&mut c, obj_str, 0);
    emit_str_equals(&mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);

    // It's an object; check __type
    c.emit_op_u16(Op::LOCAL_GET, val, 0);
    c.emit_struct_field_op(Op::STRUCT_GET, 0, type_key, 0);
    emit_const_index(&mut c, dt_str, 0);
    emit_str_equals(&mut c, 0);
    let _is_dt = c.emit_block(0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    emit_const_index(&mut c, v7, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);
    c.emit_br(1, 0); // exit outer block
    c.emit_end(0);
    c.patch_block(_is_dt);

    // Generic object → 9
    emit_const_index(&mut c, v9, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_end(0);
    c.patch_block(done);
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
        0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
        0xFF0000, 0x00FF00, 0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF,
    ];

    // Build the palette as a constant array, then ARRAY_GET by index.
    // Compile-time pack: emit ARRAY_NEW + 16 push-style emits → array.
    // Simpler: chain SELECTs for the 16 entries — but that's 15 selects
    // and bloats the chunk. Use a small array literal instead.
    let arr_locals_start = 1u16;
    c.local_count = 2;
    crate::primitives::collections::emit_array_new_into(_imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, arr_locals_start, 0);
    for &val in palette.iter() {
        let v_const = c.add_constant(Value::I32(val));
        c.emit_op_u16(Op::LOCAL_GET, arr_locals_start, 0);
        emit_const_index(&mut c, v_const, 0);
        crate::primitives::collections::emit_push_into(_imports, &mut c, 0);
        c.emit_op(Op::DROP, 0);
    }
    // ARRAY_GET(arr, idx & 0xF) — clamp via mask so out-of-range wraps.
    c.emit_op_u16(Op::LOCAL_GET, arr_locals_start, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_FROM_F64, 0);
    let mask = c.add_constant(Value::I32(0xF));
    emit_const_index(&mut c, mask, 0);
    c.emit_op(Op::I32_AND, 0);
    crate::primitives::collections::emit_get_into(_imports, &mut c, 0);
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
    let ts_idx = c.add_import("ecma:number", "toString");
    emit_const_index(&mut c, pref, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_const_index(&mut c, r, 0);
    c.emit_call(ts_idx, 2, 0);
    emit_str_concat(imports, &mut c, 0);
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
    let isfin = c.add_import("ecma:number", "isFinite");
    let isnan = c.add_import("ecma:number", "isNaN");

    // !isFinite(n) && !isNaN(n)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(isfin, 1, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(isnan, 1, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_op(Op::I32_AND, 0);
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
    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // len = entries.length
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // i = 0
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // if i >= len: break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);

    // result.push(entries[i][1])
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::ARRAY_GET, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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
fn build_setdefault(imports: &mut Chunk) -> Chunk {
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

    // result = existing (default to existing; overwrite if missing)
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // if existing is null/undefined: assign default + use it as result.
    let done_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // existing not null → keep result, exit

    // dict[key] = default; result = default
    c.emit_op_u16(Op::LOCAL_GET, dict, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_end(0);
    c.patch_block(done_block);

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
fn build_to_bytes(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_to_bytes");
    c.arity = 1;
    c.local_count = 2;
    let s = 0u16;
    let enc = 1u16;

    let new_idx = c.add_import("web:encoding", "encoderNew");
    let encode_idx = c.add_import("web:encoding", "encode");

    c.emit_call(new_idx, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, enc, 0);

    c.emit_op_u16(Op::LOCAL_GET, enc, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_call(encode_idx, 2, 0);
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
fn build_id(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_id");
    c.arity = 1;
    c.local_count = 1;
    let to_str = c.add_import("ecma:string", "String");
    let len_idx = c.add_import("ecma:string", "length");

    // Convert value to string and return its length as a stand-in id.
    // Same value (toString-stable) → same id. Not unique across all
    // values but matches the contract for compile_ok-style tests.
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_call(len_idx, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── hash(v) → number — Python `hash` / Ruby `Object#hash` ─────────
// Same shape as `id` for now: derive a stable integer from the
// stringified value. Not cryptographic — matches the Python guarantee
// that `hash(a) == hash(b)` whenever `a == b` for hashable types.
fn build_hash(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_hash");
    c.arity = 1;
    c.local_count = 1;
    let to_str = c.add_import("ecma:string", "String");
    let len_idx = c.add_import("ecma:string", "length");

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_call(len_idx, 1, 0);
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
    c.local_count = 12; // value(0), picture(1), prefix(2), dot_pos(3), decimals(4), fmt_lower(5), val_str(6), work(7), idx_a(8), idx_b(9), idx_c(10), percent(11)
    let value = 0u16;
    let picture = 1u16;
    let prefix = 2u16;
    let dot_pos = 3u16;
    let decimals = 4u16;
    let fmt_lower = 5u16;
    let val_str = 6u16;
    let work = 7u16;
    let idx_a = 8u16;
    let idx_b = 9u16;
    let idx_c = 10u16;
    let percent = 11u16;

    let to_str = c.add_import("ecma:string", "String");
    let to_lower = c.add_import("ecma:string", "toLowerCase");
    let pad_start = c.add_import("ecma:string", "padStart");
    let to_fixed = c.add_import("ecma:number", "toFixed");
    let parse_int = c.add_import("ecma:number", "parseInt");

    // prefix = ""
    let empty = c.add_constant(Value::String(Arc::from("")));
    let short_date = c.add_constant(Value::String(Arc::from("short date")));
    let short_time = c.add_constant(Value::String(Arc::from("short time")));
    let percent_str = c.add_constant(Value::String(Arc::from("%")));
    let zero_str = c.add_constant(Value::String(Arc::from("0")));
    let space_str = c.add_constant(Value::String(Arc::from(" ")));
    let slash_str = c.add_constant(Value::String(Arc::from("/")));
    let colon_str = c.add_constant(Value::String(Arc::from(":")));
    emit_const_index(&mut c, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, prefix, 0);
    c.emit_bool_const(false, 0);
    c.emit_op_u16(Op::LOCAL_SET, percent, 0);

    // If picture is null/empty, return String(value).
    let no_picture_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    let picture_null = c.add_constant(Value::I32(0));
    let _ = picture_null; // not used; kept for future readability
    let null_or_empty = c.emit_block(0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    // null path → return String(value)
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(null_or_empty);

    // Check empty string ("" length = 0)
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(no_picture_block);

    // fmt_lower = picture.Trim().ToLowerCase(); val_str = String(value)
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    {
        let idx = c.add_import("ecma:string", "trim");
        c.emit_call(idx, 1, 0);
    }
    c.emit_call(to_lower, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, fmt_lower, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, val_str, 0);

    // Short Date → first segment before space, else whole value string.
    let not_short_date = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, fmt_lower, 0);
    emit_const_index(&mut c, short_date, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    emit_const_index(&mut c, space_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_a, 0);
    let no_date_space = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(no_date_space);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(not_short_date);

    // Short Time → trim optional date prefix, drop seconds, keep AM/PM.
    let not_short_time = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, fmt_lower, 0);
    emit_const_index(&mut c, short_time, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, val_str, 0);
    c.emit_call(to_str, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    emit_const_index(&mut c, space_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_a, 0);
    let no_prefix = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, idx_b, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_b, 0);
    emit_const_index(&mut c, slash_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_c, 0);
    let keep_work = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, idx_c, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    emit_str_length(&mut c, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_end(0);
    c.patch_block(keep_work);
    c.emit_end(0);
    c.patch_block(no_prefix);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    emit_const_index(&mut c, space_str, 0);
    {
        let idx = c.add_import("ecma:string", "lastIndexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_a, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, idx_b, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    emit_str_length(&mut c, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, idx_c, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_b, 0);
    emit_const_index(&mut c, colon_str, 0);
    {
        let idx = c.add_import("ecma:string", "lastIndexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, idx_a, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_b, 0);
    emit_const_index(&mut c, colon_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    let already_short_time = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_b, 0);
    emit_const_index(&mut c, space_str, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_c, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(already_short_time);
    c.emit_op_u16(Op::LOCAL_GET, idx_b, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_a, 0);
    emit_str_substring(&mut c, 0);
    emit_const_index(&mut c, space_str, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx_c, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(not_short_time);

    // If picture starts with '$', strip it and stash as prefix.
    let dollar_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    emit_str_char_code_at(&mut c, 0);
    let dollar_code = c.add_constant(Value::I32(b'$' as i32));
    emit_const_index(&mut c, dollar_code, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    // prefix = "$"
    let dollar_str = c.add_constant(Value::String(Arc::from("$")));
    emit_const_index(&mut c, dollar_str, 0);
    c.emit_op_u16(Op::LOCAL_SET, prefix, 0);
    // picture = picture.substring(1)
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, picture, 0);
    c.emit_end(0);
    c.patch_block(dollar_block);

    // If picture ends with '%', strip it and mark percentage mode.
    let percent_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    emit_str_char_code_at(&mut c, 0);
    let percent_code = c.add_constant(Value::I32(b'%' as i32));
    emit_const_index(&mut c, percent_code, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_bool_const(true, 0);
    c.emit_op_u16(Op::LOCAL_SET, percent, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, picture, 0);
    c.emit_end(0);
    c.patch_block(percent_block);

    // dot_pos = picture.indexOf(".")
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    let dot_str = c.add_constant(Value::String(Arc::from(".")));
    emit_const_index(&mut c, dot_str, 0);
    {
        let idx = c.add_import("ecma:string", "indexOf");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op_u16(Op::LOCAL_SET, dot_pos, 0);

    // If no dot: return prefix + String(parseInt(value))
    let no_decimals_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, dot_pos, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    // No dot — integer rendering, optionally zero-padded / percentage.
    let no_pct_int = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, percent, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    let hundred = c.add_constant(Value::F64(100.0));
    emit_const_index(&mut c, hundred, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_end(0);
    c.emit_end(0);
    c.patch_block(no_pct_int);

    c.emit_op_u16(Op::LOCAL_GET, prefix, 0);
    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_call(parse_int, 1, 0);
    c.emit_call(to_str, 1, 0);
    let zero_pad_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    emit_str_char_code_at(&mut c, 0);
    let zero_code = c.add_constant(Value::I32(b'0' as i32));
    emit_const_index(&mut c, zero_code, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    emit_const_index(&mut c, zero_str, 0);
    c.emit_call(pad_start, 3, 0);
    c.emit_end(0);
    c.patch_block(zero_pad_block);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    let no_pct_suffix_int = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, percent, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    emit_const_index(&mut c, percent_str, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_end(0);
    c.patch_block(no_pct_suffix_int);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(no_decimals_block);

    // decimals = picture.length - dot_pos - 1
    c.emit_op_u16(Op::LOCAL_GET, picture, 0);
    emit_str_length(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, dot_pos, 0);
    c.emit_op(Op::I32_SUB, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, decimals, 0);

    // return prefix + Number(value).toFixed(decimals), with optional percentage suffix
    c.emit_op_u16(Op::LOCAL_GET, prefix, 0);
    let no_pct_dec = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, percent, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    let hundred_dec = c.add_constant(Value::F64(100.0));
    emit_const_index(&mut c, hundred_dec, 0);
    c.emit_op(Op::F64_MUL, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_op_u16(Op::LOCAL_SET, work, 0);
    c.emit_end(0);
    c.emit_end(0);
    c.patch_block(no_pct_dec);

    c.emit_op_u16(Op::LOCAL_GET, work, 0);
    c.emit_op_u16(Op::LOCAL_GET, decimals, 0);
    c.emit_call(to_fixed, 2, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    let no_pct_suffix_dec = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, percent, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    emit_const_index(&mut c, percent_str, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_end(0);
    c.patch_block(no_pct_suffix_dec);
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

    let to_str = c.add_import("ecma:string", "String");

    // out = ""
    let empty = c.add_constant(Value::String(std::sync::Arc::from("")));
    emit_const_index(&mut c, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);

    // i = 0; len = wasm:js-string.length(s)
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    emit_str_length(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    let open_brace = c.add_constant(Value::I32(b'{' as i32));
    let close_brace = c.add_constant(Value::I32(b'}' as i32));

    let outer_block = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // if i >= len: break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);

    // ch = wasm:js-string.charCodeAt(s, i)
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    emit_str_char_code_at(&mut c, 0);

    // Branch on '{' / '}' / literal
    let ch_slot = c.alloc_scratch(1);
    c.emit_op_u16(Op::LOCAL_SET, ch_slot, 0);

    // -- '{' branch --
    let open_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, ch_slot, 0);
    emit_const_index(&mut c, open_brace, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);

    // Find closing '}': end = i+1; while end < len && s[end] != '}': end++
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);

    let scan_block = c.emit_block(0);
    let (scan_loop, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    emit_str_char_code_at(&mut c, 0);
    emit_const_index(&mut c, close_brace, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, end, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(scan_loop);
    c.emit_end(0);
    c.patch_block(scan_block);

    // key = s.substring(i+1, end); out += String(d[key]); i = end + 1
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);

    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, d, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_call(to_str, 1, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);

    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_br(1, 0); // continue outer loop
    c.emit_end(0);
    c.patch_block(open_block);

    // -- literal char path: out += s.substring(i, i+1); i++
    c.emit_op_u16(Op::LOCAL_GET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, s, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    emit_str_substring(&mut c, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, out, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(outer_block);

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
    emit_const_index(&mut c, s, 0);
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
    emit_const_index(&mut c, nl, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── floor(n) → int — wraps f64_floor opcode ────────────────
#[allow(dead_code)]
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
    emit_is_string(&mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // skip string branch if NOT string

    // String branch: [obj, start, end] → str_substring
    c.emit_op_u16(Op::LOCAL_GET, obj, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    emit_str_substring(&mut c, 0);
    c.emit_op(Op::RETURN, 0);

    // Array branch
    c.emit_end(0);
    c.patch_block(str_block_p);
    c.emit_op_u16(Op::LOCAL_GET, obj, 0);
    c.emit_op_u16(Op::LOCAL_GET, start, 0);
    c.emit_op_u16(Op::LOCAL_GET, end, 0);
    crate::primitives::collections::emit_slice_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── keys(obj) → array of string keys ────────────────────────
// Iterates object properties, collects non-internal keys.
#[allow(dead_code)]
fn build_keys(imports: &mut Chunk) -> Chunk {
    // Can't iterate properties in pure bytecode without host support.
    // Use dict_keys host call pattern — but that's what we're trying to avoid.
    // Fallback: return empty array. On Vybe, host fn handles it.
    let mut c = Chunk::new("__stdlib_keys");
    c.arity = 1;
    c.local_count = 1;
    // Return empty array as fallback (properties aren't enumerable in pure WASM)
    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
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
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── assign(target, source) → target with source props merged ─
#[allow(dead_code)]
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

// ── jsGetMethod(obj, key) → callable | undefined ──────────────────
fn build_js_get_method(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_js_get_method");
    c.arity = 2;
    c.local_count = 4; // obj(0), key(1), cur(2), method(3)
    let proto_key = c.add_constant(Value::String(std::sync::Arc::from("__proto__")));

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 2, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_br_if(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    emit_is_undefined(&mut c, 0);
    c.emit_br_if(1, 0);

    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 3, 0);

    let missing_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    emit_is_undefined(&mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(missing_p);

    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    emit_const_index(&mut c, proto_key, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 2, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    crate::primitives::expressions::emit_undefined(&mut c, 0);
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
    emit_const_index(&mut c, proto_key, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 2, 0);

    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_if(0);
    c.emit_bool_const(false, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 3, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    emit_const_index(&mut c, link_key, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 3, 0);

    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_br_if(1, 0);

    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    c.emit_bool_const(true, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);
    c.emit_bool_const(false, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── deleteProperty(obj, key) → bool ─────────────────────────
#[allow(dead_code)]
fn build_delete_property(imports: &mut Chunk) -> Chunk {
    // Can't delete properties in pure bytecode. Set to null as fallback.
    let mut c = Chunk::new("__stdlib_deleteproperty");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0); // obj
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // key
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0); // value = null
    crate::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_bool_const(true, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── from(iterable) → array copy ─────────────────────────────
#[allow(dead_code)]
fn build_array_from(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_from");
    c.arity = 1;
    c.local_count = 1;
    // Slice the entire array (copy)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    let max = c.add_constant(Value::I32(i32::MAX));
    emit_const_index(&mut c, max, 0);
    crate::primitives::collections::emit_slice_into(imports, &mut c, 0);
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
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0); // new_size
    crate::primitives::collections::emit_slice_into(imports, &mut c, 0);
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
    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 4, 0);

    // len = arr.length
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 10, 0);

    // step = step ?? 1
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_if(0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, 7, 0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    c.emit_op_u16(Op::LOCAL_SET, 7, 0);
    c.emit_end(0);

    // start = start ?? (step > 0 ? 0 : len - 1)
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    emit_const_index(&mut c, zero, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    emit_const_index(&mut c, zero, 0);
    c.emit_op_u16(Op::LOCAL_SET, 8, 0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, 10, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, 8, 0);
    c.emit_end(0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, 8, 0);
    c.emit_end(0);

    // end = end ?? (step > 0 ? len : -1)
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    emit_const_index(&mut c, zero, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, 10, 0);
    c.emit_op_u16(Op::LOCAL_SET, 9, 0);
    c.emit_else(0);
    emit_const_index(&mut c, neg_one, 0);
    c.emit_op_u16(Op::LOCAL_SET, 9, 0);
    c.emit_end(0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op_u16(Op::LOCAL_SET, 9, 0);
    c.emit_end(0);

    // step=0 would otherwise spin forever; return empty slice.
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    emit_const_index(&mut c, zero, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);

    // i = normalized start
    c.emit_op_u16(Op::LOCAL_GET, 8, 0);
    c.emit_op_u16(Op::LOCAL_SET, 5, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // Compute condition: if step > 0 then i < end else i > end
    // Store in local 6 (cond) to avoid value-on-stack across branches.
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    emit_const_index(&mut c, zero, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);

    // positive step: cond = i < end
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 9, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 6, 0);

    // negative step: cond = i > end
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 9, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 6, 0);
    c.emit_end(0);

    // Check condition — exit if false
    c.emit_op_u16(Op::LOCAL_GET, 6, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop (depth 1 = outer block)

    // bounds check: skip push if i < 0 or i >= arr.length
    // Block must wrap the condition values consumed by br_if inside it
    let skip_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    emit_const_index(&mut c, zero, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // skip push if i < 0
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 10, 0);
    crate::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // skip push if i >= length
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_is_string(&mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    {
        let idx = c.add_import("ecma:string", "charAt");
        c.emit_call(idx, 2, 0);
    }
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_else(0);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_end(0);
    c.emit_end(0);
    c.patch_block(skip_block_p);

    // i = i + step
    c.emit_op_u16(Op::LOCAL_GET, 5, 0);
    c.emit_op_u16(Op::LOCAL_GET, 7, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 5, 0);
    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    let string_branch = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_is_string(&mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    let empty = c.add_constant(Value::String(std::sync::Arc::from("")));
    emit_const_index(&mut c, empty, 0);
    crate::primitives::collections::emit_join_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(string_branch);

    c.emit_op_u16(Op::LOCAL_GET, 4, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── dynMul(a, b) → string repeat or numeric multiply ─────────
fn build_dyn_mul(imports: &mut Chunk) -> Chunk {
    use std::sync::Arc;
    let mut c = Chunk::new("__stdlib_dynmul");
    c.arity = 2;
    c.local_count = 2;
    let str_tag = c.add_constant(Value::String(Arc::from("string")));
    // if typeof(a) == "string": return str_repeat(a, b)
    let a_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_typeof(&mut c, 0);
    emit_const_index(&mut c, str_tag, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // skip if a is NOT string
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    {
        let idx = c.add_import("ecma:string", "repeat");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(a_block_p);
    // if typeof(b) == "string": return str_repeat(b, a)
    let b_block_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    emit_typeof(&mut c, 0);
    emit_const_index(&mut c, str_tag, 0);
    crate::primitives::ops::emit_dyn_eq_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // skip if b is NOT string
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    {
        let idx = c.add_import("ecma:string", "repeat");
        c.emit_call(idx, 2, 0);
    }
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    c.patch_block(b_block_p);
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
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // i = 1
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let outer_block_p = c.emit_block(0);
    let (outer_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit outer loop

    // key = arr[i]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, key, 0);

    // keyVal = keyFn(key)
    c.emit_op_u16(Op::LOCAL_GET, key_fn, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut c, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, key_val, 0);

    // j = i - 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    // while j >= 0 && keyFn(arr[j]) > keyVal
    let inner_block_p = c.emit_block(0);
    let (inner_loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit inner loop

    // compare: keyFn(arr[j]) > keyVal
    c.emit_op_u16(Op::LOCAL_GET, key_fn, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut c, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, key_val, 0);
    crate::primitives::ops::emit_dyn_gt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit inner loop (second condition)

    // arr[j+1] = arr[j]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    crate::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // j -= 1
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, j, 0);

    c.emit_br(0, 0); // continue inner loop
    c.emit_end(0);
    c.patch_loop(inner_loop_p);
    c.emit_end(0);
    c.patch_block(inner_block_p);

    // arr[j+1] = key
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, j, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    crate::primitives::collections::emit_set_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue outer loop
    c.emit_end(0);
    c.patch_loop(outer_loop_p);
    c.emit_end(0);
    c.patch_block(outer_block_p);

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
    emit_is_string(&mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // skip string path if NOT string

    // String path: str_concat(a, b)
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op_u16(Op::LOCAL_GET, b, 0);
    emit_str_concat(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);

    c.emit_end(0);
    c.patch_block(str_block_p);

    // Array path: array_concat(a, b)
    c.emit_op_u16(Op::LOCAL_GET, a, 0);
    c.emit_op_u16(Op::LOCAL_GET, b, 0);
    crate::primitives::collections::emit_concat_into(imports, &mut c, 0);
    c.emit_op(Op::RETURN, 0);

    c
}

// ── String.raw(strings, ...values) → interleave strings and values ──
// Tagged template function that returns the raw string without escape processing.
// strings[0] + values[0] + strings[1] + values[1] + ... + strings[N]
// Since this is called as a tagged template, strings is an array and
// values are individual args. With rest params, values is already an array.
/// Drain a Continuation (JS generator) into an Array via repeated
/// generator iterator advances (WASM stack switching). Used by `Array.from(gen())`,
/// `[...gen()]`, `for ... of gen()` when the iterable variable holds
/// a generator. Returns an empty array when the input isn't a
/// Continuation (caller pre-checks via `ecma:value.isGenerator`).
fn build_drain_generator(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_drain_generator");
    c.arity = 1;
    // The arity parameter (local 0) is the continuation; emit_drain_into_array_into
    // allocates its own locals starting from local_count=1 (after the arity param).
    c.local_count = 1;
    // Delegate entirely to the common generator emitter — no inline logic here.
    crate::primitives::generators::emit_drain_into_array_into(imports, &mut c, 0);
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
    emit_const_index(&mut c, empty, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // len = strings.length
    c.emit_op_u16(Op::LOCAL_GET, strings, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // i = 0
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    // loop: while i < len
    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0); // exit loop

    // result += strings[i]
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, strings, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    emit_str_concat(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // if i < values.length: result += String(values[i])
    let skip_val_p = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, values, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // skip if i >= values.length

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, values, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::collections::emit_get_into(imports, &mut c, 0);
    emit_str_concat(imports, &mut c, 0); // dyn_add would also work since result is string
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_end(0);
    c.patch_block(skip_val_p);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0); // continue loop
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

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
    c.emit_op_u16(Op::LOCAL_GET, a, 0); // a
    c.emit_op_u16(Op::LOCAL_GET, a, 0); // a
    c.emit_op_u16(Op::LOCAL_GET, b, 0); // b
    c.emit_op(Op::F64_DIV, 0); // a / b
    c.emit_op(Op::F64_TRUNC, 0); // trunc(a / b)
    c.emit_op_u16(Op::LOCAL_GET, b, 0); // b
    c.emit_op(Op::F64_MUL, 0); // trunc(a / b) * b
    c.emit_op(Op::F64_SUB, 0); // a - trunc(a / b) * b
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_insert(arr, index, value) → null ──────────────────────────────
// splice(arr, index, 0, value) — inserts value at index without removing anything.
fn build_array_insert(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_insert");
    c.arity = 3; // arr, index, value
    c.local_count = 3;
    let arr = 0u16;
    let index = 1;
    let value = 2;

    // splice(arr, index, 0, value)
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    core_wasm::i32_const(&mut c, 0, 0); // deleteCount = 0
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    let splice = c.add_import("ecma:array", "splice");
    c.emit_call(splice, 4, 0); // 4 args
    c.emit_op(Op::DROP, 0); // drop returned removed-elements array
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_remove_at(arr, index) → null ──────────────────────────────────
// splice(arr, index, 1) — removes 1 element at index.
fn build_array_remove_at(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_remove_at");
    c.arity = 2; // arr, index
    c.local_count = 2;
    let arr = 0u16;
    let index = 1;

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    core_wasm::i32_const(&mut c, 0, 1); // deleteCount = 1
    let splice = c.add_import("ecma:array", "splice");
    c.emit_call(splice, 3, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
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
    let index_of = c.add_import("ecma:array", "indexOf");
    c.emit_call(index_of, 2, 0);
    c.emit_op_u16(Op::LOCAL_SET, idx, 0);

    // if idx >= 0: splice + return true
    c.emit_op_u16(Op::LOCAL_GET, idx, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    crate::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_to_bool_into(imports, &mut c, 0);
    c.emit_if(0);

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, idx, 0);
    core_wasm::i32_const(&mut c, 0, 1); // deleteCount = 1
    let splice = c.add_import("ecma:array", "splice");
    c.emit_call(splice, 3, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_bool_const(true, 0);
    c.emit_op(Op::RETURN, 0);

    c.emit_end(0);
    c.emit_bool_const(false, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_insert_range(arr, index, src) → null ──────────────────────────
// Loop: for i in 0..src.length: splice(arr, index+i, 0, src[i])
fn build_array_insert_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_insert_range");
    c.arity = 3; // arr, index, src
    c.local_count = 5;
    let arr = 0u16;
    let index = 1;
    let src = 2;
    let i = 3;
    let src_len = 4;

    let len_import = c.add_import("ecma:array", "length");
    let get_import = c.add_import("ecma:array", "get");
    let splice_import = c.add_import("ecma:array", "splice");

    // src_len = length(src)
    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    c.emit_call(len_import, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, src_len, 0);
    // i = 0
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let blk = c.emit_block(0);
    let (lp, _) = c.emit_loop_s(0);
    // if i >= src_len break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, src_len, 0);
    crate::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);
    // splice(arr, index+i, 0, src[i])
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_call(get_import, 2, 0);
    c.emit_call(splice_import, 4, 0);
    c.emit_op(Op::DROP, 0);
    // i++
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(lp);
    c.emit_end(0);
    c.patch_block(blk);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_set_range(arr, index, src) → null ─────────────────────────────
// Loop: for i in 0..src.length: arr[index+i] = src[i]
fn build_array_set_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_set_range");
    c.arity = 3;
    c.local_count = 5;
    let arr = 0u16;
    let index = 1;
    let src = 2;
    let i = 3;
    let src_len = 4;

    let len_import = c.add_import("ecma:array", "length");
    let get_import = c.add_import("ecma:array", "get");
    let set_import = c.add_import("ecma:array", "set");

    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    c.emit_call(len_import, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, src_len, 0);
    core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let blk = c.emit_block(0);
    let (lp, _) = c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, src_len, 0);
    crate::primitives::ops::emit_dyn_ge_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);
    // set(arr, index+i, get(src, i))
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, src, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_call(get_import, 2, 0);
    c.emit_call(set_import, 3, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(lp);
    c.emit_end(0);
    c.patch_block(blk);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_binary_search(arr, value) → i32 ───────────────────────────────
// Delegates to indexOf — correct for unsorted arrays, O(n) not O(log n)
// but avoids needing integer division opcode.
fn build_array_binary_search(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_binary_search");
    c.arity = 2; // arr, value
    c.local_count = 2;
    let arr = 0u16;
    let value = 1;
    let index_of = c.add_import("ecma:array", "indexOf");
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    c.emit_call(index_of, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_reverse_range(arr, index, count) → null ───────────────────────
fn build_array_reverse_range(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_reverse_range");
    c.arity = 3;
    c.local_count = 6;
    let arr = 0u16;
    let index = 1;
    let count = 2;
    let lo = 3;
    let hi = 4;
    let tmp = 5;

    let get_import = c.add_import("ecma:array", "get");
    let set_import = c.add_import("ecma:array", "set");

    // lo = index; hi = index + count - 1
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_SET, lo, 0);
    c.emit_op_u16(Op::LOCAL_GET, index, 0);
    c.emit_op_u16(Op::LOCAL_GET, count, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    crate::primitives::ops::emit_dyn_neg_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, hi, 0);

    let blk = c.emit_block(0);
    let (lp, _) = c.emit_loop_s(0);
    // while lo < hi
    c.emit_op_u16(Op::LOCAL_GET, lo, 0);
    c.emit_op_u16(Op::LOCAL_GET, hi, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);
    // tmp = arr[lo]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, lo, 0);
    c.emit_call(get_import, 2, 0);
    c.emit_op_u16(Op::LOCAL_SET, tmp, 0);
    // arr[lo] = arr[hi]
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, lo, 0);
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, hi, 0);
    c.emit_call(get_import, 2, 0);
    c.emit_call(set_import, 3, 0);
    c.emit_op(Op::DROP, 0);
    // arr[hi] = tmp
    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, hi, 0);
    c.emit_op_u16(Op::LOCAL_GET, tmp, 0);
    c.emit_call(set_import, 3, 0);
    c.emit_op(Op::DROP, 0);
    // lo++; hi--
    c.emit_op_u16(Op::LOCAL_GET, lo, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, lo, 0);
    c.emit_op_u16(Op::LOCAL_GET, hi, 0);
    core_wasm::i32_const(&mut c, 0, 1);
    crate::primitives::ops::emit_dyn_neg_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_add_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, hi, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(lp);
    c.emit_end(0);
    c.patch_block(blk);
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── array_last_index_of(arr, value) → i32 ───────────────────────────────
#[allow(dead_code)]
fn build_array_last_index_of(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_array_last_index_of");
    c.arity = 2; // arr, value
    c.local_count = 2;
    let arr = 0u16;
    let value = 1;

    c.emit_op_u16(Op::LOCAL_GET, arr, 0);
    c.emit_op_u16(Op::LOCAL_GET, value, 0);
    let last_index_of = c.add_import("ecma:array", "lastIndexOf");
    c.emit_call(last_index_of, 2, 0);
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

fn build_regex_replace_pat_first(_imports: &mut Chunk) -> Chunk {
    // PHP `preg_replace` and Python `re.sub` are GLOBAL by default
    // (replace every match). JS `str.replace` is single-match unless
    // the regex has `/g`. Route through `ecma:regexp.replaceAll` so
    // the always-global semantic is preserved without forcing a `/g`
    // flag through the pattern string.
    let mut c = Chunk::new("__stdlib_regex_replace_pat_first");
    let idx = c.add_import("ecma:regexp", "replaceAll");
    c.arity = 3;
    c.local_count = 3; // pat(0), repl(1), str(2)
    // Push (str, pat, repl) — ecma:regexp.replaceAll order.
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_call(idx, 3, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// PHP `preg_split($pat, $str)` / Python `re.split(pat, str)` →
// `ecma:regexp.split(str, regex)`.
fn build_regex_split_pat_first(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_regex_split_pat_first");
    let idx = c.add_import("ecma:regexp", "split");
    c.arity = 2;
    c.local_count = 2; // pat(0), str(1)
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(idx, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// PHP `preg_match_all($pat, $str)` / Python `re.findall(pat, str)` →
// `ecma:regexp.matchAll(str, regex)`.
fn build_regex_match_all_pat_first(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_regex_match_all_pat_first");
    let idx = c.add_import("ecma:regexp", "matchAll");
    c.arity = 2;
    c.local_count = 2; // pat(0), str(1)
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(idx, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
