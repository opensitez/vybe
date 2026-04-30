//! Bundle stdlib into compiled output.
//!
//! Appends stdlib chunks to program. Emits a preamble in the script chunk
//! that creates function refs and stores them in globals (`__vybe_*`).
//!
//! Call sites do: `global_get "__vybe_range"` + `call_ref 3`
//!
//! On Vybe VM, `register_all` overwrites these globals with host fn objects.
//! On any other runtime, the globals hold the bundled stdlib function refs.
//! One binary, works everywhere, no unresolvable imports.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use crate::emitter::stdlib::build_stdlib;
use std::sync::Arc;

/// Mapping from stdlib chunk name to the global name used at call sites.
const MAPPINGS: &[(&str, &str)] = &[
    ("__stdlib_range",      "__vybe_range"),
    ("__stdlib_sorted",     "__vybe_sorted"),
    ("__stdlib_sort_in_place", "__vybe_sort_in_place"),
    ("__stdlib_sort_with_comparator", "__vybe_sort_with_comparator"),
    ("__stdlib_sort_by_key", "__vybe_sort_by_key"),
    ("__stdlib_reversed",   "__vybe_reversed"),
    ("__stdlib_enumerate",  "__vybe_enumerate"),
    ("__stdlib_zip",        "__vybe_zip"),
    ("__stdlib_sum",        "__vybe_sum"),
    ("__stdlib_min",        "__vybe_min"),
    ("__stdlib_max",        "__vybe_max"),
    ("__stdlib_pyany",      "__vybe_pyany"),
    ("__stdlib_pyall",      "__vybe_pyall"),
    ("__stdlib_compact",    "__vybe_compact"),
    ("__stdlib_uniq",       "__vybe_uniq"),
    ("__stdlib_minmax",     "__vybe_minmax"),
    ("__stdlib_isempty",    "__vybe_isempty"),
    ("__stdlib_pymap",      "__vybe_pymap"),
    ("__stdlib_pyfilter",   "__vybe_pyfilter"),
    ("__stdlib_pynext",     "__vybe_pynext"),
    ("__stdlib_rand_choice","__vybe_rand_choice"),
    ("__stdlib_rand_shuffle","__vybe_rand_shuffle"),
    ("__stdlib_rand_sample","__vybe_rand_sample"),
    ("__stdlib_rotate",     "__vybe_rotate"),
    ("__stdlib_array_copy", "__vybe_array_copy"),
    ("__stdlib_pow",        "__vybe_pow"),
    ("__stdlib_sin",        "__vybe_sin"),
    ("__stdlib_cos",        "__vybe_cos"),
    ("__stdlib_tan",        "__vybe_tan"),
    ("__stdlib_asin",       "__vybe_asin"),
    ("__stdlib_acos",       "__vybe_acos"),
    ("__stdlib_atan",       "__vybe_atan"),
    ("__stdlib_atan2",      "__vybe_atan2"),
    ("__stdlib_log",        "__vybe_log"),
    ("__stdlib_log10",      "__vybe_log10"),
    ("__stdlib_exp",        "__vybe_exp"),
    ("__stdlib_sinh",       "__vybe_sinh"),
    ("__stdlib_cosh",       "__vybe_cosh"),
    ("__stdlib_tanh",       "__vybe_tanh"),
    ("__stdlib_sign",       "__vybe_sign"),
    ("__stdlib_clamp",      "__vybe_clamp"),
    ("__stdlib_tostring",   "__vybe_tostring"),
    ("__stdlib_string_is_null_or_empty", "__vybe_string_is_null_or_empty"),
    ("__stdlib_string_is_null_or_whitespace", "__vybe_string_is_null_or_whitespace"),
    ("__stdlib_str_insert", "__vybe_str_insert"),
    ("__stdlib_str_remove_start", "__vybe_str_remove_start"),
    ("__stdlib_str_remove_range", "__vybe_str_remove_range"),
    ("__stdlib_count",      "__vybe_count"),
    ("__stdlib_isnumeric",  "__vybe_isnumeric"),
    ("__stdlib_val",        "__vybe_val"),
    ("__stdlib_cchar",      "__vybe_cchar"),
    ("__stdlib_iif",        "__vybe_iif"),
    ("__stdlib_rgb",        "__vybe_rgb"),
    ("__stdlib_qbcolor",    "__vybe_qbcolor"),
    ("__stdlib_isobject",   "__vybe_isobject"),
    ("__stdlib_isdate",     "__vybe_isdate"),
    ("__stdlib_vartype",    "__vybe_vartype"),
    ("__stdlib_newline",    "__vybe_newline"),
    ("__stdlib_encoding",   "__vybe_encoding"),
    ("__stdlib_dict_values_from_entries", "__vybe_dict_values_from_entries"),
    ("__stdlib_has_value",  "__vybe_has_value"),
    ("__stdlib_invert",     "__vybe_invert"),
    ("__stdlib_setdefault", "__vybe_setdefault"),
    ("__stdlib_to_bytes",   "__vybe_to_bytes"),
    ("__stdlib_id",         "__vybe_id"),
    ("__stdlib_hash",       "__vybe_hash"),
    ("__stdlib_vb_format",  "__vybe_vb_format"),
    ("__stdlib_transform_values", "__vybe_transform_values"),
    ("__stdlib_transform_keys",   "__vybe_transform_keys"),
    ("__stdlib_php_inc",    "__vybe_php_inc"),
    ("__stdlib_php_dec",    "__vybe_php_dec"),
    ("__stdlib_format_map", "__vybe_format_map"),
    ("__stdlib_pyhex",      "__vybe_pyhex"),
    ("__stdlib_pyoct",      "__vybe_pyoct"),
    ("__stdlib_pybin",      "__vybe_pybin"),
    ("__stdlib_isinf",      "__vybe_isinf"),
    ("__stdlib_callable",   "__vybe_callable"),
    ("__stdlib_splice",     "__vybe_splice"),
    ("__stdlib_floor",      "__vybe_floor"),
    ("__stdlib_slice",      "__vybe_slice"),
    ("__stdlib_keys",       "__vybe_keys"),
    ("__stdlib_hasproperty","__vybe_hasproperty"),
    ("__stdlib_assign",     "__vybe_assign"),
    ("__stdlib_instanceof", "__vybe_instanceof"),
    ("__stdlib_js_get_method", "__vybe_js_get_method"),
    ("__stdlib_js_instanceof", "__vybe_js_instanceof"),
    ("__stdlib_deleteproperty","__vybe_deleteproperty"),
    ("__stdlib_from",       "__vybe_from"),
    ("__stdlib_redim",      "__vybe_redim"),
    ("__stdlib_slicestep",  "__vybe_slicestep"),
    ("__stdlib_dynmul",     "__vybe_dynmul"),
    ("__stdlib_concat",     "__vybe_concat"),
    ("__stdlib_string_raw", "__vybe_string_raw"),
    ("__stdlib_drain_generator", "__vybe_drain_generator"),
    ("__stdlib_fmod",               "__vybe_fmod"),
    ("__stdlib_array_insert",       "__vybe_array_insert"),
    ("__stdlib_array_remove_at",    "__vybe_array_remove_at"),
    ("__stdlib_array_remove_value", "__vybe_array_remove_value"),
    ("__stdlib_array_insert_range",  "__vybe_array_insert_range"),
    ("__stdlib_array_set_range",     "__vybe_array_set_range"),
    ("__stdlib_array_binary_search", "__vybe_array_binary_search"),
    ("__stdlib_array_reverse_range", "__vybe_array_reverse_range"),
    ("__stdlib_array_last_index_of", "__vybe_array_last_index_of"),
    ("__stdlib_sprintf",             "__vybe_sprintf"),
    ("__stdlib_to_primitive",        "__vybe_to_primitive"),
    // PHP loop-heavy array helpers — must match the build_polyfill
    // push order in stdlib.rs::build_stdlib (MAPPINGS index = chunk index).
    ("__stdlib_php_array_pad",           "__vybe_php_array_pad"),
    ("__stdlib_php_array_chunk",         "__vybe_php_array_chunk"),
    ("__stdlib_php_array_flip",          "__vybe_php_array_flip"),
    ("__stdlib_php_array_combine",       "__vybe_php_array_combine"),
    ("__stdlib_php_array_diff",          "__vybe_php_array_diff"),
    ("__stdlib_php_array_intersect",     "__vybe_php_array_intersect"),
    ("__stdlib_php_array_diff_assoc",    "__vybe_php_array_diff_assoc"),
    ("__stdlib_php_array_intersect_key", "__vybe_php_array_intersect_key"),
    ("__stdlib_php_array_replace",       "__vybe_php_array_replace"),
    ("__stdlib_php_array_count_values",  "__vybe_php_array_count_values"),
    ("__stdlib_php_array_column",        "__vybe_php_array_column"),
    ("__stdlib_php_array_key_first",     "__vybe_php_array_key_first"),
    ("__stdlib_php_array_key_last",      "__vybe_php_array_key_last"),
    ("__stdlib_php_asort",               "__vybe_php_asort"),
    ("__stdlib_php_arsort",              "__vybe_php_arsort"),
    ("__stdlib_php_ksort",               "__vybe_php_ksort"),
    ("__stdlib_php_krsort",              "__vybe_php_krsort"),
    ("__stdlib_php_uasort",              "__vybe_php_uasort"),
    ("__stdlib_php_uksort",              "__vybe_php_uksort"),
    ("__stdlib_php_checkdate",           "__vybe_php_checkdate"),
    ("__stdlib_php_getdate",             "__vybe_php_getdate"),
    ("__stdlib_php_ucwords",             "__vybe_php_ucwords"),
    ("__stdlib_php_str_split",           "__vybe_php_str_split"),
    ("__stdlib_php_str_pad",             "__vybe_php_str_pad"),
    ("__stdlib_php_substr_count",        "__vybe_php_substr_count"),
    ("__stdlib_php_substr_replace",      "__vybe_php_substr_replace"),
    ("__stdlib_php_str_ireplace",        "__vybe_php_str_ireplace"),
    ("__stdlib_php_str_word_count",      "__vybe_php_str_word_count"),
    ("__stdlib_php_strstr",              "__vybe_php_strstr"),
    ("__stdlib_php_stristr",             "__vybe_php_stristr"),
    ("__stdlib_php_urlencode",           "__vybe_php_urlencode"),
    ("__stdlib_php_rawurlencode",        "__vybe_php_rawurlencode"),
    ("__stdlib_php_urldecode",           "__vybe_php_urldecode"),
    ("__stdlib_php_bin2hex",             "__vybe_php_bin2hex"),
    ("__stdlib_php_hex2bin",             "__vybe_php_hex2bin"),
    ("__stdlib_php_chunk_split",         "__vybe_php_chunk_split"),
    ("__stdlib_php_wordwrap",            "__vybe_php_wordwrap"),
    ("__stdlib_php_number_format",       "__vybe_php_number_format"),
    ("__stdlib_php_str_replace",         "__vybe_php_str_replace"),
    ("__stdlib_php_ctype_alpha",         "__vybe_php_ctype_alpha"),
    ("__stdlib_php_ctype_digit",         "__vybe_php_ctype_digit"),
    ("__stdlib_php_ctype_alnum",         "__vybe_php_ctype_alnum"),
    ("__stdlib_php_ctype_space",         "__vybe_php_ctype_space"),
    ("__stdlib_php_ctype_upper",         "__vybe_php_ctype_upper"),
    ("__stdlib_php_ctype_lower",         "__vybe_php_ctype_lower"),
    ("__stdlib_php_ctype_xdigit",        "__vybe_php_ctype_xdigit"),
    ("__stdlib_php_ctype_punct",         "__vybe_php_ctype_punct"),
    ("__stdlib_php_ctype_print",         "__vybe_php_ctype_print"),
    ("__stdlib_php_ctype_cntrl",         "__vybe_php_ctype_cntrl"),
    ("__stdlib_php_min",                 "__vybe_php_min"),
    ("__stdlib_php_max",                 "__vybe_php_max"),
    ("__stdlib_php_decbin",              "__vybe_php_decbin"),
    ("__stdlib_php_decoct",              "__vybe_php_decoct"),
    ("__stdlib_php_dechex",              "__vybe_php_dechex"),
    ("__stdlib_php_base_convert",        "__vybe_php_base_convert"),
    ("__stdlib_dir_read",           "__vybe_dir_read"),
    ("__stdlib_dir_close",          "__vybe_dir_close"),
    ("__stdlib_dir",                "__vybe_dir"),
    ("__stdlib_file",               "__vybe_file"),
    ("__stdlib_filemtime",          "__vybe_filemtime"),
    ("__stdlib_file_exists",        "__vybe_file_exists"),
    ("__stdlib_is_file",            "__vybe_is_file"),
    ("__stdlib_is_dir",             "__vybe_is_dir"),
    ("__stdlib_filesize",           "__vybe_filesize"),
    ("__stdlib_unlink",             "__vybe_unlink"),
    // Regex pattern-first adapters (PHP preg_*, Python re.*, VB Regex.*)
    ("__stdlib_regex_replace_pat_first",   "__vybe_regex_replace_pat_first"),
    ("__stdlib_regex_split_pat_first",     "__vybe_regex_split_pat_first"),
    ("__stdlib_regex_match_all_pat_first", "__vybe_regex_match_all_pat_first"),
];

/// Emit the stdlib preamble at the START of a script chunk.
/// This must be called BEFORE any user code is emitted.
///
/// The preamble emits, for each stdlib function:
///   `if globals[name] is null { globals[name] = ref_func(stdlib_chunk) }`
///
/// This is the polyfill pattern: if the host (Vybe VM) has already populated
/// `__vybe_*` globals with optimized native fns BEFORE running the script,
/// the preamble leaves those alone. On non-Vybe runtimes the globals start
/// null and the preamble installs the bundled stdlib bytecode chunks.
pub fn emit_stdlib_preamble(script: &mut Chunk, stdlib_base: usize) {
    for (i, &(_chunk_name, global_name)) in MAPPINGS.iter().enumerate() {
        let ci = stdlib_base + i;
        let name_c = script.add_constant(Value::String(Arc::from(global_name)));

        // Check if global is already set: global_get + ref_is_null
        script.emit_op_u16(Op::GLOBAL_GET, name_c, 0);
        script.emit_op(Op::REF_IS_NULL, 0);
        // br_if_false skip — if NOT null, skip the install
        let skip = script.emit_jump(Op::BR_IF_FALSE, 0);

        // Install: ref_func + global_set + drop
        script.emit_op_u16(Op::REF_FUNC, ci as u16, 0);
        script.emit(0, 0); // 0 upvalues
        script.emit_op_u16(Op::GLOBAL_SET, name_c, 0);
        script.emit_op(Op::DROP, 0);

        script.patch_jump(skip);
    }
}

/// Append stdlib chunks to program chunks. Call AFTER compilation is done.
pub fn append_stdlib_chunks(program_chunks: &mut Vec<Chunk>) {
    let stdlib = {
        let (first, _rest) = program_chunks.split_at_mut(1);
        build_stdlib(&mut first[0])
    };
    program_chunks.extend(stdlib.chunks);
}

/// Emit a call to a stdlib/vybe function.
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

/// Append stdlib chunks to a compiled program and register them as global_inits.
/// Call this at the END of compilation, after all user chunks are finalized.
/// The script chunk (chunks[0]) gets global_inits with RefFunc entries.
pub fn finalize_with_stdlib(chunks: &mut Vec<Chunk>) {
    use vybe_bytecode::chunk::{GlobalInit, ConstExpr};

    let stdlib_base = chunks.len();
    // Build stdlib chunks with their imports registered on `chunks[0]`
    // (the module-level imports section — single per WASM module).
    // Same dependency surface as user code → stdlib becomes a true
    // cross-runtime polyfill: on Vybe the chunks are swapped for
    // native handlers; on v8 their `ecma:array.*` imports resolve
    // to native `Array.prototype.*`; on wasmtime the Phase-C polyfill
    // supplies the imports.
    let stdlib = {
        let (first, _rest) = chunks.split_at_mut(1);
        crate::emitter::stdlib::build_stdlib(&mut first[0])
    };

    for (i, &(chunk_name, global_name)) in MAPPINGS.iter().enumerate() {
        if stdlib.exports.iter().any(|&n| n == chunk_name) {
            chunks[0].global_inits.push(GlobalInit {
                name: global_name.to_string(),
                init: ConstExpr::RefFunc(stdlib_base + i),
            });
        }
    }
    chunks.extend(stdlib.chunks);
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
