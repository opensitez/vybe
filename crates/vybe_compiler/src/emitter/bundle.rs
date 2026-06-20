//! Link runtime helper chunks into compiled output.
//!
//! Appends helper chunks to program. Emits a preamble in the script chunk
//! that creates function refs and stores them in globals (`__vybe_*`).
//!
//! Call sites do: `global_get "__vybe_range"` + `call_ref 3`
//!
//! On Vybe VM, `register_all` overwrites these globals with host fn objects.
//! On any other runtime, the globals hold the bundled helper function refs.
//! One binary, works everywhere, no unresolvable imports.

use crate::emitter::runtime_helpers::build_runtime_helpers;
use std::collections::BTreeSet;
use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

/// Mapping from helper chunk name to the global name used at call sites.
const MAPPINGS: &[(&str, &str)] = &[
    ("__stdlib_range", "__vybe_range"),
    ("__stdlib_sorted", "__vybe_sorted"),
    ("__stdlib_sort_in_place", "__vybe_sort_in_place"),
    (
        "__stdlib_sort_with_comparator",
        "__vybe_sort_with_comparator",
    ),
    // C qsort/bsearch moved to platforms/libc/stdlib_runtime.rs (libc surface);
    // no longer part of the cross-language helper bundle.
    ("__stdlib_sort_by_key", "__vybe_sort_by_key"),
    ("__stdlib_reversed", "__vybe_reversed"),
    ("__stdlib_enumerate", "__vybe_enumerate"),
    ("__stdlib_zip", "__vybe_zip"),
    ("__stdlib_sum", "__vybe_sum"),
    ("__stdlib_min", "__vybe_min"),
    ("__stdlib_max", "__vybe_max"),
    ("__stdlib_pyany", "__vybe_pyany"),
    ("__stdlib_pyall", "__vybe_pyall"),
    ("__stdlib_compact", "__vybe_compact"),
    ("__stdlib_uniq", "__vybe_uniq"),
    ("__stdlib_minmax", "__vybe_minmax"),
    ("__stdlib_isempty", "__vybe_isempty"),
    ("__stdlib_pymap", "__vybe_pymap"),
    ("__stdlib_pyfilter", "__vybe_pyfilter"),
    ("__stdlib_pyiter", "__vybe_pyiter"),
    ("__stdlib_pynext", "__vybe_pynext"),
    ("__stdlib_rand_choice", "__vybe_rand_choice"),
    ("__stdlib_rand_shuffle", "__vybe_rand_shuffle"),
    ("__stdlib_rand_sample", "__vybe_rand_sample"),
    ("__stdlib_rotate", "__vybe_rotate"),
    ("__stdlib_array_copy", "__vybe_array_copy"),
    ("__stdlib_sin", "__vybe_sin"),
    ("__stdlib_cos", "__vybe_cos"),
    ("__stdlib_tan", "__vybe_tan"),
    ("__stdlib_asin", "__vybe_asin"),
    ("__stdlib_acos", "__vybe_acos"),
    ("__stdlib_atan", "__vybe_atan"),
    ("__stdlib_atan2", "__vybe_atan2"),
    ("__stdlib_log", "__vybe_log"),
    ("__stdlib_log10", "__vybe_log10"),
    ("__stdlib_exp", "__vybe_exp"),
    ("__stdlib_sinh", "__vybe_sinh"),
    ("__stdlib_cosh", "__vybe_cosh"),
    ("__stdlib_tanh", "__vybe_tanh"),
    ("__stdlib_sign", "__vybe_sign"),
    ("__stdlib_clamp", "__vybe_clamp"),
    ("__stdlib_tostring", "__vybe_tostring"),
    (
        "__stdlib_string_is_null_or_empty",
        "__vybe_string_is_null_or_empty",
    ),
    (
        "__stdlib_string_is_null_or_whitespace",
        "__vybe_string_is_null_or_whitespace",
    ),
    ("__stdlib_str_insert", "__vybe_str_insert"),
    ("__stdlib_str_remove_start", "__vybe_str_remove_start"),
    ("__stdlib_str_remove_range", "__vybe_str_remove_range"),
    ("__stdlib_pascal_set_include", "__vybe_pascal_set_include"),
    ("__stdlib_pascal_set_exclude", "__vybe_pascal_set_exclude"),
    ("__stdlib_pascal_set_union", "__vybe_pascal_set_union"),
    (
        "__stdlib_pascal_set_intersection",
        "__vybe_pascal_set_intersection",
    ),
    (
        "__stdlib_pascal_set_difference",
        "__vybe_pascal_set_difference",
    ),
    ("__stdlib_pascal_set_contains", "__vybe_pascal_set_contains"),
    ("__stdlib_pascal_write", "__vybe_pascal_write"),
    ("__stdlib_pascal_writeln", "__vybe_pascal_writeln"),
    ("__stdlib_pascal_str_insert", "__vybe_pascal_str_insert"),
    (
        "__stdlib_pascal_str_remove_range",
        "__vybe_pascal_str_remove_range",
    ),
    ("__stdlib_count", "__vybe_count"),
    ("__stdlib_isnumeric", "__vybe_isnumeric"),
    ("__stdlib_val", "__vybe_val"),
    ("__stdlib_cchar", "__vybe_cchar"),
    ("__stdlib_iif", "__vybe_iif"),
    ("__stdlib_rgb", "__vybe_rgb"),
    ("__stdlib_qbcolor", "__vybe_qbcolor"),
    ("__stdlib_isobject", "__vybe_isobject"),
    ("__stdlib_isdate", "__vybe_isdate"),
    ("__stdlib_vartype", "__vybe_vartype"),
    ("__stdlib_newline", "__vybe_newline"),
    ("__stdlib_encoding", "__vybe_encoding"),
    (
        "__stdlib_dict_values_from_entries",
        "__vybe_dict_values_from_entries",
    ),
    ("__stdlib_has_value", "__vybe_has_value"),
    ("__stdlib_invert", "__vybe_invert"),
    ("__stdlib_setdefault", "__vybe_setdefault"),
    ("__stdlib_to_bytes", "__vybe_to_bytes"),
    ("__stdlib_id", "__vybe_id"),
    ("__stdlib_hash", "__vybe_hash"),
    ("__stdlib_vb_format", "__vybe_vb_format"),
    (
        "__stdlib_dotnet_numeric_format",
        "__vybe_dotnet_numeric_format",
    ),
    ("__stdlib_transform_values", "__vybe_transform_values"),
    ("__stdlib_transform_keys", "__vybe_transform_keys"),
    ("__stdlib_format_map", "__vybe_format_map"),
    ("__stdlib_pyhex", "__vybe_pyhex"),
    ("__stdlib_pyoct", "__vybe_pyoct"),
    ("__stdlib_pybin", "__vybe_pybin"),
    ("__stdlib_isinf", "__vybe_isinf"),
    ("__stdlib_callable", "__vybe_callable"),
    ("__stdlib_splice", "__vybe_splice"),
    ("__stdlib_slice", "__vybe_slice"),
    ("__stdlib_hasproperty", "__vybe_hasproperty"),
    ("__stdlib_instanceof", "__vybe_instanceof"),
    ("__stdlib_js_get_method", "__vybe_js_get_method"),
    ("__stdlib_js_instanceof", "__vybe_js_instanceof"),
    ("__stdlib_redim", "__vybe_redim"),
    ("__stdlib_slicestep", "__vybe_slicestep"),
    ("__stdlib_dynmul", "__vybe_dynmul"),
    ("__stdlib_concat", "__vybe_concat"),
    ("__stdlib_string_raw", "__vybe_string_raw"),
    ("__stdlib_drain_generator", "__vybe_drain_generator"),
    ("__stdlib_fmod", "__vybe_fmod"),
    ("__stdlib_array_insert", "__vybe_array_insert"),
    ("__stdlib_array_remove_at", "__vybe_array_remove_at"),
    ("__stdlib_array_remove_value", "__vybe_array_remove_value"),
    ("__stdlib_array_insert_range", "__vybe_array_insert_range"),
    ("__stdlib_array_set_range", "__vybe_array_set_range"),
    ("__stdlib_array_binary_search", "__vybe_array_binary_search"),
    ("__stdlib_array_reverse_range", "__vybe_array_reverse_range"),
    ("__stdlib_sprintf", "__vybe_sprintf"),
    ("__stdlib_generator_next", "__vybe_generator_next"),
    ("__stdlib_generator_self", "__vybe_generator_self"),
    ("__stdlib_iter_drain", "__vybe_iter_drain"),
    // PHP runtime helpers — all inline opcode emitters under
    // `emitter/php/<cat>_adapter.rs`. Reached through `common:php.*`
    // dispatch arms; no `__vybe_*` global indirection.
    // PHP filesystem helpers — inline opcode emitters under
    // `emitter/php/filesystem_adapter.rs`. Reached via `common:php.*`
    // dispatch arms; no `__vybe_*` global indirection.
    // Regex pattern-first adapters (PHP preg_*, Python re.*, VB Regex.*)
    (
        "__stdlib_regex_replace_pat_first",
        "__ecma_regexp_replace_pat_first",
    ),
    (
        "__stdlib_regex_split_pat_first",
        "__ecma_regexp_split_pat_first",
    ),
    (
        "__stdlib_regex_match_all_pat_first",
        "__ecma_regexp_match_all_pat_first",
    ),
];

/// Emit the runtime-helper preamble at the START of a script chunk.
/// This must be called BEFORE any user code is emitted.
///
/// The preamble emits, for each helper function:
///   `if globals[name] is null { globals[name] = ref_func(stdlib_chunk) }`
///
/// This is the polyfill pattern: if the host (Vybe VM) has already populated
/// `__vybe_*` globals with optimized native fns BEFORE running the script,
/// the preamble leaves those alone. On non-Vybe runtimes the globals start
/// null and the preamble installs the bundled helper bytecode chunks.
pub fn emit_runtime_helper_preamble(script: &mut Chunk, stdlib_base: usize) {
    for (i, &(_chunk_name, global_name)) in MAPPINGS.iter().enumerate() {
        let ci = stdlib_base + i;
        let name_c = script.add_constant(Value::String(Arc::from(global_name)));

        // Check if global is already set: global_get + ref_is_null
        script.emit_op_u16(Op::GLOBAL_GET, name_c, 0);
        script.emit_op(Op::REF_IS_NULL, 0);
        let install_block = script.emit_block(0);
        script.emit_op(Op::I32_EQZ, 0);
        script.emit_br_if(0, 0);

        // Install: ref_func + global_set + drop
        script.emit_op_u16(Op::REF_FUNC, ci as u16, 0);
        script.emit(0, 0); // 0 upvalues
        script.emit_op_u16(Op::GLOBAL_SET, name_c, 0);
        script.emit_op(Op::DROP, 0);

        script.emit_end(0);
        script.patch_block(install_block);
    }
}

/// Append every runtime helper chunk to program chunks. Call AFTER compilation is done.
pub fn append_runtime_helper_chunks(program_chunks: &mut Vec<Chunk>) {
    let helpers = {
        let (first, _rest) = program_chunks.split_at_mut(1);
        build_runtime_helpers(&mut first[0])
    };
    program_chunks.extend(helpers.chunks);
}

/// Emit a call to a runtime-helper/vybe function.
/// IMPORTANT: push the function ref FIRST (via global_get), then push args, then call_ref.
/// The call convention is: [func_ref, arg0, arg1, ...] on stack.
///
/// Usage in compiler:
///   emit_call_push_func(chunk, "__vybe_range", 0);  // push func ref
///   compile_expr(arg0);                               // push start
///   compile_expr(arg1);                               // push stop
///   compile_expr(arg2);                               // push step
///   emit_call_invoke(chunk, 3, 0);                    // call_ref 3
pub fn emit_call_push_func(chunk: &mut Chunk, global_name: &str, line: u32) {
    let name_c = chunk.add_constant(Value::String(Arc::from(global_name)));
    chunk.emit_op_u16(Op::GLOBAL_GET, name_c, line);
}

/// Emit call_ref after func + args are on stack.
pub fn emit_call_invoke(chunk: &mut Chunk, argc: u8, line: u32) {
    chunk.emit_op_u8(Op::CALL_REF, argc, line);
}

/// Append referenced runtime helper chunks and register them as global_inits.
/// Call this at the END of compilation, after all user chunks are finalized.
/// The script chunk (chunks[0]) gets global_inits with RefFunc entries.
pub fn finalize_with_runtime_helpers(chunks: &mut Vec<Chunk>) {
    use vybe_bytecode::chunk::{ConstExpr, GlobalInit};

    let mut requested = referenced_helper_exports(chunks);
    add_helper_dependencies(&mut requested);
    if requested.is_empty() {
        return;
    }

    let helper_base = chunks.len();
    // Build helper chunks with their imports registered on `chunks[0]`
    // (the module-level imports section — single per WASM module).
    // Same dependency surface as user code → helpers become true
    // cross-runtime polyfill: on Vybe the chunks are swapped for
    // native handlers; on v8 their `ecma:array.*` imports resolve
    // to native `Array.prototype.*`; on wasmtime the Phase-C polyfill
    // supplies the imports.
    let helpers = {
        let (first, _rest) = chunks.split_at_mut(1);
        let exports = ordered_helper_exports(&requested);
        crate::emitter::runtime_helpers::build_runtime_helpers_for_exports(&mut first[0], &exports)
    };

    for (i, &chunk_name) in helpers.exports.iter().enumerate() {
        if let Some(global_name) = helper_global_for_export(chunk_name) {
            chunks[0].global_inits.push(GlobalInit {
                name: global_name.to_string(),
                init: ConstExpr::RefFunc(helper_base + i),
            });
        }
    }
    chunks.extend(helpers.chunks);
}

/// Convenience: emit func_ref push + args already on stack + call_ref.
/// Args must already be on stack BEFORE calling this.
/// This function inserts the func ref below the args using a temp local approach.
/// Actually — we can't easily insert below stack items without a local.
/// Instead, callers should use emit_call_push_func + args + emit_call_invoke.
///
/// For backward compatibility, this emits global_get + call_ref, but callers
/// MUST push args AFTER calling this (which means this only works for 0-arg calls).
/// For multi-arg calls, use the push_func/invoke pair.
pub fn emit_call(chunk: &mut Chunk, global_name: &str, argc: u8, line: u32) {
    let name_c = chunk.add_constant(Value::String(Arc::from(global_name)));
    chunk.emit_op_u16(Op::GLOBAL_GET, name_c, line);
    chunk.emit_op_u8(Op::CALL_REF, argc, line);
}

fn referenced_helper_exports(chunks: &[Chunk]) -> BTreeSet<&'static str> {
    let mut exports = BTreeSet::new();
    for chunk in chunks {
        for constant in &chunk.constants {
            if let Value::String(name) = constant {
                if let Some(export) = helper_export_for_global(name.as_ref()) {
                    exports.insert(export);
                }
            }
        }
    }
    // Any program with a generator needs the continuation protocol helpers.
    // `attach_continuation_protocols` (VM) wires `__vybe_generator_next` /
    // `__vybe_generator_self` onto every continuation as its real `next` /
    // `[Symbol.iterator]` methods — `__stdlib_generator_next` is a pure
    // bytecode driver (`emit_next` → spec `resume`). Generators drive via
    // inline `emit_next`, so nothing *references* these globals by name; force
    // them in so a generator is a first-class iterator (`it.next` is readable
    // and callable) for generic for-of / `Iterator.from` / spread paths.
    if chunks.iter().any(|chunk| chunk.is_generator) {
        exports.insert("__stdlib_generator_next");
        exports.insert("__stdlib_generator_self");
    }
    exports
}

fn add_helper_dependencies(exports: &mut BTreeSet<&'static str>) {
    loop {
        let before = exports.len();
        for export in exports.clone() {
            for dep in helper_export_dependencies(export) {
                exports.insert(dep);
            }
        }
        if exports.len() == before {
            break;
        }
    }
}

fn helper_export_dependencies(export: &str) -> &'static [&'static str] {
    match export {
        "__stdlib_minmax" => &["__stdlib_min", "__stdlib_max"],
        "__stdlib_pynext" => &["__stdlib_iter_drain"],
        "__stdlib_rand_sample" => &["__stdlib_rand_shuffle"],
        "__stdlib_rotate" => &["__stdlib_fmod"],
        _ => &[],
    }
}

fn ordered_helper_exports(exports: &BTreeSet<&'static str>) -> Vec<&'static str> {
    MAPPINGS
        .iter()
        .filter_map(|&(chunk_name, _global_name)| {
            exports.contains(chunk_name).then_some(chunk_name)
        })
        .collect()
}

fn helper_export_for_global(global_name: &str) -> Option<&'static str> {
    MAPPINGS.iter().find_map(|&(chunk_name, mapped_global)| {
        (mapped_global == global_name).then_some(chunk_name)
    })
}

fn helper_global_for_export(export: &str) -> Option<&'static str> {
    MAPPINGS
        .iter()
        .find_map(|&(chunk_name, global_name)| (chunk_name == export).then_some(global_name))
}
