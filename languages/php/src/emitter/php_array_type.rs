//! The PHP array as one Wasm GC struct type.
//!
//! `$PhpArray` is reserved through the compiler's type registry, so an
//! allocation carries its rtt and `is_array` is `ref.test $PhpArray`.
//!
//! - field 0 holds the array's ECMA storage: an `ecma:array` list, or an
//!   `ecma:map` once a string key has been stored;
//! - field 1 is the copy-on-write mark: null while this wrapper owns its
//!   storage, non-null once another wrapper shares it;
//! - field 2 is the internal pointer used by `current`/`next`/`reset`.
//!
//! Every ECMA operation runs on the storage, never on the wrapper.

use vybe_compiler::primitives::classes;
use vybe_runtime::opcode::heaptype::{HeapType, HT_EXTERN};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// PHP arrays are `$PhpArray` values. While this is false the PHP frontend
/// keeps raw ECMA storage as the array value and none of the leaves below is
/// emitted.
pub const ENABLED: bool = false;

pub const TYPE_NAME: &str = "PhpArray";
pub const FIELD_STORE: u32 = 0;
pub const FIELD_SHARED: u32 = 1;
pub const FIELD_CURSOR: u32 = 2;
const FIELD_NAMES: [&str; 3] = ["__php_array_store", "__php_array_shared", "__php_array_cursor"];

/// `$PhpArray` is the first type of a PHP module (see
/// `vybe_ast::ArrayWrapper`), so its index is fixed.
pub const TYPE_INDEX: u16 = 1;

pub fn field_names() -> Vec<String> {
    FIELD_NAMES.iter().map(|name| name.to_string()).collect()
}

/// The module's `$PhpArray` type index.
pub fn typeidx(chunks: &mut [Chunk]) -> u16 {
    let idx = classes::reserve_platform_type(chunks, &[TYPE_NAME.to_string()]);
    debug_assert_eq!(idx, TYPE_INDEX, "`$PhpArray` must be the module's first type");
    idx
}

/// Stack: [storage] -> [array]. A new owner of `storage`, unshared, cursor at
/// the first element.
pub fn emit_wrap(chunks: &mut [Chunk], current: usize, line: u32) {
    let ty = typeidx(chunks);
    let chunk = &mut chunks[current];
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_struct_new(ty, 3, line);
}

/// Stack: [value] -> [i32].
pub fn emit_is_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let ty = typeidx(chunks);
    chunks[current].emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(ty as u32), line);
}

/// Stack: [array] -> [storage], for a value whose producer is `$PhpArray`.
pub fn emit_store_of_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let ty = typeidx(chunks);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_STORE, line);
}

/// Stack: [value] -> [storage]. A value that is not a `$PhpArray` is already
/// storage and passes through unchanged.
pub fn emit_store_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let ty = typeidx(chunks);
    emit_store_of_in(&mut chunks[current], ty, line);
}

/// [`emit_store_of`] given the module's `$PhpArray` type index.
pub fn emit_store_of_in(chunk: &mut Chunk, ty: u16, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
    chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(ty as u32), line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_STORE, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_end(line);
}

/// Push the copy-on-write mark a shared wrapper carries in field 1.
fn emit_shared_mark(chunk: &mut Chunk, line: u32) {
    chunk.emit_bool_const(true, line);
}

/// Stack: [value] -> [storage] for a write. A shared `$PhpArray` first gets
/// storage of its own; any other value is already storage.
pub fn emit_store_of_mut(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let ty = typeidx(chunks);
    let separate = ensure_separate_helper(chunks);
    let value = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
        chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(ty as u32), line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_SHARED, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_STORE, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    }
    emit_helper_call(chunks, current, separate, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_end(line);
}

/// Stack: [value] -> [value]. Assignment, argument passing and return by
/// value. A `$PhpArray` is shared in O(1): both wrappers are marked and the
/// new one names the same storage. Raw array storage is copied once and
/// wrapped; any other value passes through.
pub fn emit_share(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let ty = typeidx(chunks);
    let reshare = ensure_reshare_helper(chunks);
    let value = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
        chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(ty as u32), line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        emit_shared_mark(chunk, line);
        chunk.emit_struct_field_op(Op::STRUCT_SET, ty as u32, FIELD_SHARED, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_STORE, line);
        emit_shared_mark(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_CURSOR, line);
        chunk.emit_struct_new(ty, 3, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        call_import(chunk, "php:array", "isArray", 1, line);
        call_import(chunk, "wasm:js-number", "toI32", 1, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        call_import(chunk, "php:array", "copy", 1, line);
        chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
    }
    emit_helper_call(chunks, current, reshare, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    }
    emit_wrap(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn call_import(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

fn emit_helper_call(chunks: &mut Vec<Chunk>, current: usize, helper: usize, line: u32) {
    if vybe_compiler::primitives::functions::try_emit_direct_helper_call(
        chunks, current, helper, 1, line,
    ) {
        return;
    }
    let arg = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, arg, line);
    chunk.emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arg, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
}

/// Stack: [value] -> [result], calling the arity-1 `helper` by reference.
pub fn emit_helper_call_in(chunk: &mut Chunk, helper: usize, line: u32) {
    let arg = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, arg, line);
    chunk.emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arg, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
}

/// The index of `__php_array_deep_raw`, built on first use.
pub fn ensure_deep_raw(chunks: &mut Vec<Chunk>) -> usize {
    ensure_deep_raw_helper(chunks)
}

fn helper_index(chunks: &[Chunk], name: &str) -> Option<usize> {
    chunks.iter().position(|chunk| chunk.name == name)
}

/// `__php_array_separate(array) -> storage`: give a shared `$PhpArray`
/// storage of its own. The copy is the existing PHP array copy, which copies
/// nested raw arrays and keeps every object's identity; nested wrappers it
/// kept are then shared through `__php_array_reshare`.
fn ensure_separate_helper(chunks: &mut Vec<Chunk>) -> usize {
    const NAME: &str = "__php_array_separate";
    if let Some(index) = helper_index(chunks, NAME) {
        return index;
    }
    let ty = typeidx(chunks);
    let index = chunks.len();
    let mut chunk = Chunk::new(NAME);
    chunk.arity = 1;
    chunk.local_count = 2;
    chunks.push(chunk);
    let reshare = ensure_reshare_helper(chunks);
    let line = 0;
    {
        let chunk = &mut chunks[index];
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_STORE, line);
        call_import(chunk, "php:array", "copy", 1, line);
        chunk.emit_op_u16(Op::LOCAL_TEE, 1, line);
    }
    emit_helper_call(chunks, index, reshare, line);
    let chunk = &mut chunks[index];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, ty as u32, FIELD_STORE, line);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, ty as u32, FIELD_SHARED, line);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, line);
    chunk.emit_op(Op::RETURN, line);
    index
}

/// `__php_array_reshare(storage) -> null`: in storage that was just copied,
/// replace every nested `$PhpArray` by a new wrapper sharing its storage and
/// mark both, descending into nested raw arrays (fresh copies themselves).
fn ensure_reshare_helper(chunks: &mut Vec<Chunk>) -> usize {
    const NAME: &str = "__php_array_reshare";
    if let Some(index) = helper_index(chunks, NAME) {
        return index;
    }
    let ty = typeidx(chunks);
    let index = chunks.len();
    let mut chunk = Chunk::new(NAME);
    chunk.arity = 1;
    chunk.local_count = 1;
    chunks.push(chunk);
    let line = 0;
    let (keys, i, n, key, element) = {
        let chunk = &mut chunks[index];
        (
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
        )
    };
    super::array_adapter::emit_php_key_list_from_slot(chunks, index, 0, line);
    {
        let chunk = &mut chunks[index];
        chunk.emit_op_u16(Op::LOCAL_TEE, keys, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        chunk.emit_op_u16(Op::LOCAL_SET, n, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    }
    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, index, line);
    {
        let chunk = &mut chunks[index];
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op_u16(Op::LOCAL_GET, n, line);
        chunk.emit_op(Op::I32_LT_U, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, index, line);
    {
        let chunk = &mut chunks[index];
        chunk.emit_op_u16(Op::LOCAL_GET, keys, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        chunk.emit_op_u16(Op::LOCAL_GET, key, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_TEE, element, line);
        chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(ty as u32), line);
        chunk.emit_if(line);
        // The nested wrapper stays with the other holder of the storage;
        // this copy gets a new one naming the same storage.
        chunk.emit_op_u16(Op::LOCAL_GET, element, line);
        emit_shared_mark(chunk, line);
        chunk.emit_struct_field_op(Op::STRUCT_SET, ty as u32, FIELD_SHARED, line);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        chunk.emit_op_u16(Op::LOCAL_GET, key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, element, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_STORE, line);
        emit_shared_mark(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, element, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_CURSOR, line);
        chunk.emit_struct_new(ty, 3, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, element, line);
        call_import(chunk, "php:array", "isArray", 1, line);
        call_import(chunk, "wasm:js-number", "toI32", 1, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, element, line);
    }
    emit_helper_call(chunks, index, index, line);
    {
        let chunk = &mut chunks[index];
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, index, state, line);
    let chunk = &mut chunks[index];
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_op(Op::RETURN, line);
    index
}

/// How a `php.*` leaf receives one of its arguments.
#[derive(Clone, Copy, PartialEq)]
pub enum Argument {
    /// As given: the leaf stores it, or knows `$PhpArray` itself.
    Keep,
    /// The storage, for a leaf that reads the array.
    Read,
    /// The storage after separation, for a leaf that writes the array.
    Write,
    /// A fresh raw copy with every nested `$PhpArray` replaced by raw
    /// storage, for a leaf that walks nested arrays.
    Deep,
}

/// The argument policy of `leaf` at `position`.
pub fn argument(leaf: &str, position: usize) -> Argument {
    use Argument::*;
    let name = leaf.strip_prefix("php.").unwrap_or(leaf);
    match name {
        "array_read_key" | "array_read_deref" | "offset" | "quiet_index" | "array_change_key_case"
        | "array_chunk" | "array_count_values" | "array_flip" | "array_is_list"
        | "array_key_first" | "array_key_last" | "array_keys" | "array_values" | "array_rand"
        | "array_reverse" | "array_slice" | "array_sum" | "iterator_to_array" | "count" | "uniq" => {
            if position == 0 { Read } else { Keep }
        }
        "array_fill_keys" | "array_pad" => if position == 0 { Read } else { Keep },
        // The walker passes the array first.
        "in_array" | "array_key_exists" => if position == 0 { Read } else { Keep },
        "preg_grep" => if position == 1 { Read } else { Keep },
        "array_search" | "strtr" => if position == 1 { Read } else { Keep },
        "array_to_object" | "session_set_cookie_params" | "unset_capture" | "spl_arrayiterator"
        | "spl_arrayobject" | "spl_cachingiterator" | "spl_infiniteiterator"
        | "spl_iteratoriterator" | "spl_recursiveiteratoriterator" | "spl_recursivetreeiterator"
        | "spl_autoload_register" | "spl_autoload_unregister" => {
            if position == 0 { Read } else { Keep }
        }
        "array_splice_assoc_removed" => if position == 0 { Write } else { Read },
        "arrayobject_exchange" => if position == 1 { Deep } else { Keep },
        "implode" | "min" | "max" | "array_map" | "array_merge" | "array_replace" | "array_diff"
        | "array_diff_assoc" | "array_diff_key" | "array_intersect" | "array_intersect_assoc"
        | "array_intersect_key" | "array_combine" | "zip" => {
            Read
        }
        "array_filter" | "reduce_array" | "array_reduce" => if position < 2 { Read } else { Keep },
        "array_udiff" | "array_uintersect" | "array_udiff_uassoc" | "array_uintersect_uassoc"
        | "array_intersect_ukey" => Read,
        "str_replace" | "str_ireplace" => if position < 3 { Read } else { Keep },
        "preg_replace_callback" | "regex_replace_pat_first" => if position < 3 { Read } else { Keep },
        "regex_match_all_pat_first" => if position == 2 { Read } else { Keep },
        "array_column" | "http_build_query" | "serialize" | "json_encode" | "json_encode_ordered_map"
        | "var_dump_stringify" | "print_r"
        | "var_dump" | "var_export" | "array_merge_recursive" | "array_replace_recursive" => {
            if position == 0 || matches!(name, "array_merge_recursive" | "array_replace_recursive") {
                Deep
            } else {
                Keep
            }
        }
        "sort_g" | "asort" | "arsort" | "ksort" | "krsort" | "uasort" | "uksort"
        | "sort_with_comparator" | "array_walk" | "array_walk_recursive" => {
            if position == 0 { Write } else { Read }
        }
        "array_push_g" | "array_pop_g" => if position == 0 { Write } else { Keep },
        "array_multisort" => Write,
        "call_user_func_array" => if position < 2 { Read } else { Keep },
        "pdo_statement_execute" => Read,
        "dynamic_method_call" => if position == 2 { Read } else { Keep },
        "local_set" | "local_append" | "cursor_reset" | "cursor_next" | "cursor_end"
        | "cursor_prev" | "array_unshift" | "array_shift" => if position == 0 { Write } else { Keep },
        "cursor_current" | "cursor_key" | "foreach_entries" | "unpack_values" => {
            if position == 0 { Read } else { Keep }
        }
        _ => Keep,
    }
}

/// Stack: [args...] -> [args...] with each argument of `leaf` mapped by its
/// [`argument`] policy. Nothing is emitted for a leaf that keeps them all.
pub fn emit_leaf_arguments(chunks: &mut Vec<Chunk>, current: usize, leaf: &str, argc: u8, line: u32) {
    let policies: Vec<Argument> = (0..argc as usize).map(|position| argument(leaf, position)).collect();
    if policies.iter().all(|policy| *policy == Argument::Keep) {
        return;
    }
    let first = chunks[current].alloc_scratch(argc as u16);
    for position in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, first + position, line);
    }
    for (position, policy) in policies.iter().enumerate() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, first + position as u16, line);
        match policy {
            Argument::Keep => {}
            Argument::Read => emit_store_of(chunks, current, line),
            Argument::Write => emit_store_of_mut(chunks, current, line),
            Argument::Deep => {
                let deep = ensure_deep_raw_helper(chunks);
                emit_helper_call(chunks, current, deep, line);
            }
        }
    }
}

/// Stack: [value] -> [value], through `__php_array_deep_raw`.
pub fn emit_deep_raw(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let deep = ensure_deep_raw_helper(chunks);
    emit_helper_call(chunks, current, deep, line);
}

/// `__php_array_deep_raw(value) -> value`: an array value as a fresh raw
/// copy whose nested arrays are raw copies too; any other value unchanged.
fn ensure_deep_raw_helper(chunks: &mut Vec<Chunk>) -> usize {
    const NAME: &str = "__php_array_deep_raw";
    if let Some(index) = helper_index(chunks, NAME) {
        return index;
    }
    let ty = typeidx(chunks);
    let index = chunks.len();
    let mut chunk = Chunk::new(NAME);
    chunk.arity = 1;
    chunk.local_count = 1;
    chunks.push(chunk);
    let line = 0;
    let (copy, keys, i, n, key) = {
        let chunk = &mut chunks[index];
        (
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
        )
    };
    {
        let chunk = &mut chunks[index];
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(ty as u32), line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_STORE, line);
        chunk.emit_op_u16(Op::LOCAL_SET, 0, line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        call_import(chunk, "php:array", "isArray", 1, line);
        call_import(chunk, "wasm:js-number", "toI32", 1, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        chunk.emit_op(Op::RETURN, line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
        call_import(chunk, "php:array", "copy", 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, copy, line);
    }
    super::array_adapter::emit_php_key_list_from_slot(chunks, index, copy, line);
    {
        let chunk = &mut chunks[index];
        chunk.emit_op_u16(Op::LOCAL_TEE, keys, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        chunk.emit_op_u16(Op::LOCAL_SET, n, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    }
    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, index, line);
    {
        let chunk = &mut chunks[index];
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op_u16(Op::LOCAL_GET, n, line);
        chunk.emit_op(Op::I32_LT_U, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, index, line);
    {
        let chunk = &mut chunks[index];
        chunk.emit_op_u16(Op::LOCAL_GET, keys, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op_u16(Op::LOCAL_SET, key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, copy, line);
        chunk.emit_op_u16(Op::LOCAL_GET, key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, copy, line);
        chunk.emit_op_u16(Op::LOCAL_GET, key, line);
        chunk.emit_op(Op::ARRAY_GET, line);
    }
    emit_helper_call(chunks, index, index, line);
    {
        let chunk = &mut chunks[index];
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, index, state, line);
    let chunk = &mut chunks[index];
    chunk.emit_op_u16(Op::LOCAL_GET, copy, line);
    chunk.emit_op(Op::RETURN, line);
    index
}

/// Stack: [value] -> [i32]. True for a `$PhpArray` and for raw array storage.
pub fn emit_is_php_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let ty = typeidx(chunks);
    emit_is_php_array_in(&mut chunks[current], ty, line);
}

/// The module's `$PhpArray` type index while [`ENABLED`], else 0.
pub fn array_type(chunks: &mut [Chunk]) -> u16 {
    if ENABLED { typeidx(chunks) } else { 0 }
}

/// [`emit_is_php_array`] given the module's `$PhpArray` type index.
pub fn emit_is_php_array_in(chunk: &mut Chunk, ty: u16, line: u32) {
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
    chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(ty as u32), line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunk, "php:array", "isArray", 1, line);
    call_import(chunk, "wasm:js-number", "toI32", 1, line);
    chunk.emit_end(line);
}

/// Stack: [value] -> [value]. A shared `$PhpArray` first gets storage of its
/// own, so a write through this wrapper reaches no other holder.
pub fn emit_own(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let ty = typeidx(chunks);
    let separate = ensure_separate_helper(chunks);
    let value = chunks[current].alloc_scratch(1);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
        chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(ty as u32), line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, ty as u32, FIELD_SHARED, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    }
    emit_helper_call(chunks, current, separate, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
}
