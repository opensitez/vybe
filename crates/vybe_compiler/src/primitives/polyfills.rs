//! Polyfills — the linkable-chunk half of the primitives layer.
//!
//! A primitive reaches compiled code two ways. Most splice their
//! instructions straight into the call site (`emit_*`, living in the
//! topical module: `strings.rs`, `collections.rs`, `arrays.rs`, …). The
//! rest are whole functions the module has to *contain* — because the
//! body is large (`sprintf`), because it recurses, or because the
//! language needs it as a first-class value. Those are built as
//! standalone `Chunk`s (`build_*`, also in the topical modules) and
//! bundled here.
//!
//! This file is the registry that ties the two together: export name →
//! builder. `bundle.rs` walks the resulting [`RuntimeHelpers`] and, for
//! each export, binds a `__vybe_*` global to the chunk. Linking is
//! on-demand — a chunk is built only when compiled code references its
//! global — so this is a menu, not a bundled standard library.
//!
//! It also holds the polyfill factory ([`build_polyfill`]): a chunk
//! whose body is authored in a source language and compiled through the
//! normal frontend, rather than hand-emitted as opcode calls. Same
//! artifact, same registry, different way of writing the body.
//!
//! Nothing here is a *runtime* concern despite the historical name —
//! every decision in this file is made at compile time, and what it
//! produces is what a compiled module carries with it.

use vybe_runtime::Chunk;

// Generator DRIVER chunks live with the rest of the generator machinery
// in `crate::primitives::generators` — this module only registers them.
use crate::primitives::generators::{build_async_generator_next, build_generator_next};

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

    chunks.push(crate::primitives::collections::build_sorted(imports));
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
    chunks.push(crate::primitives::collections::build_sort_in_place(imports));
    exports.push("__stdlib_sort_in_place");
    chunks.push(crate::primitives::collections::build_sort_with_comparator(imports));
    exports.push("__stdlib_sort_with_comparator");
    // `__stdlib_reversed` removed — `reversed()` inlines its polymorphic loop
    // in `crate::primitives::collections::emit_reversed`.
    chunks.push(crate::primitives::collections::build_enumerate(imports));
    exports.push("__stdlib_enumerate");
    chunks.push(crate::primitives::collections::build_sum(imports));
    exports.push("__stdlib_sum");
    chunks.push(crate::primitives::collections::build_min(imports));
    exports.push("__stdlib_min");
    chunks.push(crate::primitives::collections::build_max(imports));
    exports.push("__stdlib_max");
    chunks.push(crate::primitives::collections::build_pyany(imports));
    exports.push("__stdlib_pyany");
    chunks.push(crate::primitives::collections::build_pyall(imports));
    exports.push("__stdlib_pyall");
    chunks.push(crate::primitives::collections::build_compact(imports));
    exports.push("__stdlib_compact");
    chunks.push(crate::primitives::collections::build_isempty(imports));
    exports.push("__stdlib_isempty");
    chunks.push(crate::primitives::collections::build_pymap(imports));
    exports.push("__stdlib_pymap");
    chunks.push(crate::primitives::collections::build_pyfilter(imports));
    exports.push("__stdlib_pyfilter");
    chunks.push(crate::primitives::collections::build_pyiter(imports));
    exports.push("__stdlib_pyiter");
    chunks.push(crate::primitives::collections::build_pynext(imports));
    exports.push("__stdlib_pynext");
    chunks.push(crate::primitives::arrays::build_array_copy(imports));
    exports.push("__stdlib_array_copy");
    // Math transcendentals (sin/cos/tan/…/sign/clamp) removed: dead chunks —
    // every language routes math through `Math.*` → `ecma:math:*` host fns
    // directly, so these `env`-delegating wrappers were never bundled.
    // `__stdlib_tostring` removed — `str()` / `toString` route to
    // `ecma:string.String` directly (Python via `emit_helper`, others via
    // `emit_to_string`).
    chunks.push(crate::primitives::strings::build_string_is_null_or_empty(imports));
    exports.push("__stdlib_string_is_null_or_empty");
    chunks.push(crate::primitives::strings::build_string_is_null_or_whitespace(imports));
    exports.push("__stdlib_string_is_null_or_whitespace");
    chunks.push(crate::primitives::strings::build_str_insert(imports));
    exports.push("__stdlib_str_insert");
    chunks.push(crate::primitives::strings::build_str_remove_start(imports));
    exports.push("__stdlib_str_remove_start");
    chunks.push(crate::primitives::strings::build_str_remove_range(imports));
    exports.push("__stdlib_str_remove_range");
    chunks.push(crate::primitives::sets::build_pascal_set_include(imports));
    exports.push("__stdlib_pascal_set_include");
    chunks.push(crate::primitives::sets::build_pascal_set_exclude(imports));
    exports.push("__stdlib_pascal_set_exclude");
    chunks.push(crate::primitives::sets::build_pascal_set_union(imports));
    exports.push("__stdlib_pascal_set_union");
    chunks.push(crate::primitives::sets::build_pascal_set_intersection(imports));
    exports.push("__stdlib_pascal_set_intersection");
    chunks.push(crate::primitives::sets::build_pascal_set_difference(imports));
    exports.push("__stdlib_pascal_set_difference");
    chunks.push(crate::primitives::sets::build_pascal_set_contains(imports));
    exports.push("__stdlib_pascal_set_contains");
    chunks.push(crate::primitives::strings::build_pascal_str_insert(imports));
    exports.push("__stdlib_pascal_str_insert");
    chunks.push(crate::primitives::strings::build_pascal_str_remove_range(imports));
    exports.push("__stdlib_pascal_str_remove_range");
    chunks.push(crate::primitives::convert::build_is_numeric(imports));
    exports.push("__stdlib_isnumeric");
    chunks.push(crate::primitives::convert::build_val(imports));
    exports.push("__stdlib_val");
    chunks.push(crate::primitives::control_flow::build_iif(imports));
    exports.push("__stdlib_iif");
    chunks.push(crate::primitives::gui::build_rgb(imports));
    exports.push("__stdlib_rgb");
    chunks.push(crate::primitives::gui::build_qbcolor(imports));
    exports.push("__stdlib_qbcolor");
    chunks.push(crate::primitives::reflection::build_isdate(imports));
    exports.push("__stdlib_isdate");
    chunks.push(crate::primitives::reflection::build_vartype(imports));
    exports.push("__stdlib_vartype");
    chunks.push(crate::primitives::strings::build_newline(imports));
    exports.push("__stdlib_newline");
    chunks.push(crate::primitives::dict::build_dict_values_from_entries(imports));
    exports.push("__stdlib_dict_values_from_entries");
    chunks.push(crate::primitives::dict::build_setdefault(imports));
    exports.push("__stdlib_setdefault");
    chunks.push(crate::primitives::convert::build_to_bytes(imports));
    exports.push("__stdlib_to_bytes");
    chunks.push(crate::primitives::reflection::build_id(imports));
    exports.push("__stdlib_id");
    chunks.push(crate::primitives::reflection::build_hash(imports));
    exports.push("__stdlib_hash");
    chunks.push(crate::primitives::strings::build_vb_format(imports));
    exports.push("__stdlib_vb_format");
    chunks.push(vybe_runtime::registry::platform_numeric_format_helper()
        .expect("no platform registered a numeric-format helper")(
        imports
    ));
    exports.push("__stdlib_dotnet_numeric_format");
    // PHP `$x++` / `$x--` stay in the PHP emitter path (`common:php.{inc,dec}`)
    // rather than going through bundled stdlib/polyfill helpers.
    chunks.push(crate::primitives::strings::build_format_map(imports));
    exports.push("__stdlib_format_map");
    chunks.push(crate::primitives::convert::build_pyradix(imports, "__stdlib_pyhex", "0x", 16));
    exports.push("__stdlib_pyhex");
    chunks.push(crate::primitives::convert::build_pyradix(imports, "__stdlib_pyoct", "0o", 8));
    exports.push("__stdlib_pyoct");
    chunks.push(crate::primitives::convert::build_pyradix(imports, "__stdlib_pybin", "0b", 2));
    exports.push("__stdlib_pybin");
    chunks.push(crate::primitives::math::build_isinf(imports));
    exports.push("__stdlib_isinf");
    // `__stdlib_splice` removed — no emit site references the `__vybe_splice`
    // global (the chunk was bundled but never used).
    // `__stdlib_slice` removed — slicing uses direct polymorphic `ecma:array.slice`.
    chunks.push(crate::primitives::object::build_js_get_method(imports));
    exports.push("__stdlib_js_get_method");
    chunks.push(crate::primitives::arrays::build_redim(imports));
    exports.push("__stdlib_redim");
    chunks.push(crate::primitives::slices::build_slice_step(imports));
    exports.push("__stdlib_slicestep");
    chunks.push(crate::primitives::ops::build_dyn_mul(imports));
    exports.push("__stdlib_dynmul");
    chunks.push(crate::primitives::math::build_fmod(imports));
    exports.push("__stdlib_fmod");
    chunks.push(crate::primitives::arrays::build_array_insert(imports));
    exports.push("__stdlib_array_insert");
    chunks.push(crate::primitives::arrays::build_array_remove_at(imports));
    exports.push("__stdlib_array_remove_at");
    chunks.push(crate::primitives::arrays::build_array_remove_value(imports));
    exports.push("__stdlib_array_remove_value");
    chunks.push(crate::primitives::arrays::build_array_insert_range(imports));
    exports.push("__stdlib_array_insert_range");
    chunks.push(crate::primitives::arrays::build_array_set_range(imports));
    exports.push("__stdlib_array_set_range");
    chunks.push(crate::primitives::arrays::build_array_binary_search(imports));
    exports.push("__stdlib_array_binary_search");
    chunks.push(crate::primitives::arrays::build_array_reverse_range(imports));
    exports.push("__stdlib_array_reverse_range");
    // The receiver ABI is a MODULE property and rides on the module chunk —
    // which `imports` IS. Read it before the `&mut` borrows below.
    let abi = imports.module_receiver_abi;
    chunks.push(build_generator_next(imports));
    exports.push("__stdlib_generator_next");
    chunks.push(build_async_generator_next(imports));
    exports.push("__stdlib_async_generator_next");
    chunks.push(crate::primitives::generators::build_generator_self());
    exports.push("__stdlib_generator_self");
    chunks.push(crate::primitives::generators::build_iter_drain(imports));
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

    // Regex adapters: pattern-first (PHP `preg_*`, Python `re.*`) →
    // ECMA string-first. Same Layer-3 shape as the `String.Format`
    // dotnet adapter.
    chunks.push(crate::primitives::regex::build_regex_replace_pat_first(imports));
    exports.push("__stdlib_regex_replace_pat_first");
    chunks.push(crate::primitives::regex::build_regex_split_pat_first(imports));
    exports.push("__stdlib_regex_split_pat_first");
    chunks.push(crate::primitives::regex::build_regex_match_all_pat_first(imports));
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
        "__stdlib_sorted" => crate::primitives::collections::build_sorted(imports),
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
        "__stdlib_sort_in_place" => crate::primitives::collections::build_sort_in_place(imports),
        "__stdlib_sort_with_comparator" => crate::primitives::collections::build_sort_with_comparator(imports),
        "__stdlib_enumerate" => crate::primitives::collections::build_enumerate(imports),
        "__stdlib_sum" => crate::primitives::collections::build_sum(imports),
        "__stdlib_min" => crate::primitives::collections::build_min(imports),
        "__stdlib_max" => crate::primitives::collections::build_max(imports),
        "__stdlib_pyany" => crate::primitives::collections::build_pyany(imports),
        "__stdlib_pyall" => crate::primitives::collections::build_pyall(imports),
        "__stdlib_compact" => crate::primitives::collections::build_compact(imports),
        "__stdlib_isempty" => crate::primitives::collections::build_isempty(imports),
        "__stdlib_pymap" => crate::primitives::collections::build_pymap(imports),
        "__stdlib_pyfilter" => crate::primitives::collections::build_pyfilter(imports),
        "__stdlib_pyiter" => crate::primitives::collections::build_pyiter(imports),
        "__stdlib_pynext" => crate::primitives::collections::build_pynext(imports),
        "__stdlib_array_copy" => crate::primitives::arrays::build_array_copy(imports),
        "__stdlib_tostring" => crate::primitives::strings::build_to_string(imports),
        "__stdlib_string_is_null_or_empty" => crate::primitives::strings::build_string_is_null_or_empty(imports),
        "__stdlib_string_is_null_or_whitespace" => crate::primitives::strings::build_string_is_null_or_whitespace(imports),
        "__stdlib_str_insert" => crate::primitives::strings::build_str_insert(imports),
        "__stdlib_str_remove_start" => crate::primitives::strings::build_str_remove_start(imports),
        "__stdlib_str_remove_range" => crate::primitives::strings::build_str_remove_range(imports),
        "__stdlib_pascal_set_include" => crate::primitives::sets::build_pascal_set_include(imports),
        "__stdlib_pascal_set_exclude" => crate::primitives::sets::build_pascal_set_exclude(imports),
        "__stdlib_pascal_set_union" => crate::primitives::sets::build_pascal_set_union(imports),
        "__stdlib_pascal_set_intersection" => crate::primitives::sets::build_pascal_set_intersection(imports),
        "__stdlib_pascal_set_difference" => crate::primitives::sets::build_pascal_set_difference(imports),
        "__stdlib_pascal_set_contains" => crate::primitives::sets::build_pascal_set_contains(imports),
        "__stdlib_pascal_str_insert" => crate::primitives::strings::build_pascal_str_insert(imports),
        "__stdlib_pascal_str_remove_range" => crate::primitives::strings::build_pascal_str_remove_range(imports),
        "__stdlib_isnumeric" => crate::primitives::convert::build_is_numeric(imports),
        "__stdlib_val" => crate::primitives::convert::build_val(imports),
        "__stdlib_iif" => crate::primitives::control_flow::build_iif(imports),
        "__stdlib_rgb" => crate::primitives::gui::build_rgb(imports),
        "__stdlib_qbcolor" => crate::primitives::gui::build_qbcolor(imports),
        "__stdlib_isdate" => crate::primitives::reflection::build_isdate(imports),
        "__stdlib_vartype" => crate::primitives::reflection::build_vartype(imports),
        "__stdlib_newline" => crate::primitives::strings::build_newline(imports),
        "__stdlib_dict_values_from_entries" => crate::primitives::dict::build_dict_values_from_entries(imports),
        "__stdlib_setdefault" => crate::primitives::dict::build_setdefault(imports),
        "__stdlib_to_bytes" => crate::primitives::convert::build_to_bytes(imports),
        "__stdlib_id" => crate::primitives::reflection::build_id(imports),
        "__stdlib_hash" => crate::primitives::reflection::build_hash(imports),
        "__stdlib_vb_format" => crate::primitives::strings::build_vb_format(imports),
        "__stdlib_dotnet_numeric_format" => {
            vybe_runtime::registry::platform_numeric_format_helper()
                .expect("no platform registered a numeric-format helper")(imports)
        }
        "__stdlib_format_map" => crate::primitives::strings::build_format_map(imports),
        "__stdlib_pyhex" => crate::primitives::convert::build_pyradix(imports, "__stdlib_pyhex", "0x", 16),
        "__stdlib_pyoct" => crate::primitives::convert::build_pyradix(imports, "__stdlib_pyoct", "0o", 8),
        "__stdlib_pybin" => crate::primitives::convert::build_pyradix(imports, "__stdlib_pybin", "0b", 2),
        "__stdlib_isinf" => crate::primitives::math::build_isinf(imports),
        "__stdlib_splice" => crate::primitives::slices::build_splice(imports),
        "__stdlib_slice" => crate::primitives::slices::build_slice(imports),
        "__stdlib_js_get_method" => crate::primitives::object::build_js_get_method(imports),
        "__stdlib_redim" => crate::primitives::arrays::build_redim(imports),
        "__stdlib_slicestep" => crate::primitives::slices::build_slice_step(imports),
        "__stdlib_dynmul" => crate::primitives::ops::build_dyn_mul(imports),
        "__stdlib_fmod" => crate::primitives::math::build_fmod(imports),
        "__stdlib_array_insert" => crate::primitives::arrays::build_array_insert(imports),
        "__stdlib_array_remove_at" => crate::primitives::arrays::build_array_remove_at(imports),
        "__stdlib_array_remove_value" => crate::primitives::arrays::build_array_remove_value(imports),
        "__stdlib_array_insert_range" => crate::primitives::arrays::build_array_insert_range(imports),
        "__stdlib_array_set_range" => crate::primitives::arrays::build_array_set_range(imports),
        "__stdlib_array_binary_search" => crate::primitives::arrays::build_array_binary_search(imports),
        "__stdlib_array_reverse_range" => crate::primitives::arrays::build_array_reverse_range(imports),
        "__stdlib_generator_next" => {
            let abi = imports.module_receiver_abi;
            build_generator_next(imports)
        }
        "__stdlib_async_generator_next" => {
            let abi = imports.module_receiver_abi;
            build_async_generator_next(imports)
        }
        "__stdlib_generator_self" => {
            crate::primitives::generators::build_generator_self()
        }
        "__stdlib_iter_drain" => {
            let abi = imports.module_receiver_abi;
            crate::primitives::generators::build_iter_drain(imports)
        }
        "__stdlib_regex_replace_pat_first" => {
            crate::primitives::regex::build_regex_replace_pat_first(imports)
        }
        "__stdlib_regex_split_pat_first" => {
            crate::primitives::regex::build_regex_split_pat_first(imports)
        }
        "__stdlib_regex_match_all_pat_first" => {
            crate::primitives::regex::build_regex_match_all_pat_first(imports)
        }
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

// ── Removed from the registry, recorded so they stay removed ─────────
//
// Math transcendentals: `env`-module-importing wrappers
// (`add_import("env", "sin")`) — a host dependency that shouldn't
// exist. `ecma:math` provides sin/cos/tan/asin/acos/atan/atan2/log/
// log10/exp/sinh/cosh/tanh/sign natively and every language routes
// `Math.*` → `ecma:math:*` directly, so these were dead AND pulled in a
// phantom `env` import.
//
// PHP filesystem builders: now inline opcode emitters in
// `emitter/php/filesystem_adapter.rs`, reached via `common:php.*`
// dispatch arms with no `__vybe_*` global indirection.
//
// `rest_fixed_arity(name)`: a one-entry table asserting that `sprintf`
// has 1 fixed param before its rest. That fact is DERIVED from the
// declaration everywhere it is actually needed — `lambdas.rs` and
// `classes.rs` compute it as the param count preceding the rest param
// and feed `Compiler::rest_fixed_arities`, which `calls.rs` reads to
// split individual args from packed ones. Looking a builtin's arity up
// by name was the pre-declaration workaround; it had no callers left.

// Regex adapters (pattern-first → ECMA string-first) live in
// `crate::primitives::regex` — this module only registers them.
