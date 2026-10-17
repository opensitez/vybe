//! Owned, capture-free literal replacement algorithms and PHP conversion order.
use vybe_compiler::primitives::{functions, references};
use vybe_runtime::chunk::{LocalName, LocalOrigin};
use vybe_runtime::{Chunk, Op};

fn get(c: &mut Chunk, slot: u16, line: u32) { c.emit_op_u16(Op::LOCAL_GET, slot, line); }
fn set(c: &mut Chunk, slot: u16, line: u32) { c.emit_op_u16(Op::LOCAL_SET, slot, line); }
fn call(c: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let import = c.add_import(module, name);
    c.emit_call(import, argc, line);
}
fn direct(chunks: &mut Vec<Chunk>, current: usize, target: usize, argc: u8, line: u32) {
    assert!(functions::try_emit_direct_helper_call(chunks, current, target, argc, line),
        "owned replacement helper must retain its capture-free generic ABI");
}
fn receipt(name: &str, arity: u8) -> String {
    format!("<php-replace:v1:{name}:{arity}:generic>")
}
fn owned(chunks: &[Chunk], name: &str, arity: u8) -> Option<usize> {
    let marker = receipt(name, arity);
    chunks.iter().position(|c| c.name == name && c.source_path.is_none()
        && c.arity == arity && c.result_arity == 1 && c.func_sig.is_none()
        && !c.takes_receiver && !c.is_async && !c.is_generator
        && c.capture_count == 0 && c.direct_entry_has_captures != Some(true)
        && c.local_names.iter().any(|n| n.slot == 0
            && n.origin == LocalOrigin::Compiler && n.name == marker))
}
fn append_owned(chunks: &mut Vec<Chunk>, name: &str, arity: u8) -> usize {
    let mut c = Chunk::new(name);
    c.arity = arity;
    c.local_count = u16::from(arity);
    c.local_names.push(LocalName::new(0, receipt(name, arity), LocalOrigin::Compiler));
    let index = chunks.len();
    chunks.push(c);
    index
}
fn width(c: &mut Chunk, slot: u16) {
    get(c, slot, 0);
    call(c, "wasm:js-string", "length", 1, 0);
    call(c, "wasm:js-number", "toF64", 1, 0);
}

/// Arguments 0..3 are physical subject/needle/replacement Strings. Argument
/// 3, when present, is an owned one-element numeric accumulator, never a PHP
/// place/reference. Reject nonstrings even on the empty-needle path.
fn ensure_literal(chunks: &mut Vec<Chunk>, counted: bool) -> usize {
    let (name, arity) = if counted { ("__stdlib_php_literal_replace_counted", 4) }
        else { ("__stdlib_php_literal_replace", 3) };
    if let Some(index) = owned(chunks, name, arity) { return index; }
    let current = append_owned(chunks, name, arity);
    let c = &mut chunks[current];
    for slot in 0..3 {
        get(c, slot, 0);
        call(c, "wasm:js-string", "test", 1, 0);
        if slot != 0 { c.emit_op(Op::I32_AND, 0); }
    }
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_if(0);
    super::type_guard::emit_throw_const(chunks, current, "TypeError",
        "Internal literal replacement requires physical strings", 0);
    let c = &mut chunks[current];
    c.emit_end(0);
    if counted {
        // A missing literal cannot change either the subject or accumulator.
        // Keep the physical-string guard above this: even an absent needle
        // must not make invalid raw replacement values silently acceptable.
        // Empty needles return zero here and retain the existing empty case.
        get(c, 0, 0);
        get(c, 1, 0);
        call(c, "ecma:string", "indexOf", 2, 0);
        c.emit_f64_const(0.0, 0);
        c.emit_op(Op::F64_LT, 0);
        c.emit_if(0);
        get(c, 0, 0);
        c.emit_op(Op::RETURN, 0);
        c.emit_end(0);
    }
    let n = c.alloc_scratch(1);
    let r = c.alloc_scratch(1);
    let result = c.alloc_scratch(1);
    let hits = c.alloc_scratch(1);
    let pieces = c.alloc_scratch(1);
    width(c, 1);
    set(c, n, 0);
    get(c, n, 0);
    c.emit_f64_const(0.0, 0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_if(0);
    get(c, 0, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    if counted {
        width(c, 2);
        set(c, r, 0);
        get(c, n, 0);
        get(c, r, 0);
        c.emit_op(Op::F64_NE, 0);
    }
    // Dollar substitution belongs to ECMAScript, not PHP literal replacement.
    // A split/join fallback is literal on both native and external providers.
    get(c, 2, 0);
    c.emit_string_const("$", 0);
    call(c, "ecma:string", "indexOf", 2, 0);
    c.emit_f64_const(0.0, 0);
    c.emit_op(Op::F64_LT, 0);
    if counted { c.emit_op(Op::I32_AND, 0); }
    c.emit_if(0);
    for slot in 0..3 { get(c, slot, 0); }
    call(c, "ecma:string", "replaceAll", 3, 0);
    set(c, result, 0);
    if counted {
        width(c, result);
        width(c, 0);
        c.emit_op(Op::F64_SUB, 0);
        get(c, r, 0);
        get(c, n, 0);
        c.emit_op(Op::F64_SUB, 0);
        c.emit_op(Op::F64_DIV, 0);
        set(c, hits, 0);
    }
    c.emit_else(0);
    get(c, 0, 0);
    get(c, 1, 0);
    call(c, "ecma:string", "split", 2, 0);
    c.emit_op_u16(Op::LOCAL_TEE, pieces, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    call(c, "wasm:js-number", "toF64", 1, 0);
    c.emit_f64_const(1.0, 0);
    c.emit_op(Op::F64_SUB, 0);
    c.emit_op_u16(Op::LOCAL_TEE, hits, 0);
    c.emit_f64_const(0.0, 0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_if(0);
    get(c, 0, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    get(c, pieces, 0);
    get(c, 2, 0);
    call(c, "ecma:array", "join", 2, 0);
    set(c, result, 0);
    c.emit_end(0);
    if counted {
        get(c, 3, 0);
        c.emit_i32_const(0, 0);
        get(c, 3, 0);
        c.emit_i32_const(0, 0);
        c.emit_op(Op::ARRAY_GET, 0);
        get(c, hits, 0);
        c.emit_op(Op::F64_ADD, 0);
        c.emit_op(Op::ARRAY_SET, 0);
    }
    get(c, result, 0);
    c.emit_op(Op::RETURN, 0);
    current
}

fn coerce(chunks: &mut Vec<Chunk>, current: usize, slot: u16, stringify: usize) {
    let c = &mut chunks[current];
    get(c, slot, 0);
    call(c, "wasm:js-string", "test", 1, 0);
    c.emit_if_value(0);
    get(c, slot, 0);
    c.emit_else(0);
    get(c, slot, 0);
    direct(chunks, current, stringify, 1, 0);
    let c = &mut chunks[current];
    c.emit_end(0);
    set(c, slot, 0);
}
fn read_element(chunks: &mut Vec<Chunk>, current: usize, slot: u16) {
    set(&mut chunks[current], slot, 0);
    get(&mut chunks[current], slot, 0);
    references::emit_autoderef_to_stack(chunks, current, 0);
    set(&mut chunks[current], slot, 0);
}
fn snapshot(chunks: &mut Vec<Chunk>, current: usize, input: u16, flag: u16,
    output: u16, entries: bool) {
    let c = &mut chunks[current];
    get(c, flag, 0);
    c.emit_if(0);
    get(c, input, 0);
    call(c, "ecma:object", if entries { "entries" } else { "values" }, 1, 0);
    set(c, output, 0);
    c.emit_end(0);
}

#[derive(Clone, Copy)]
struct Search {
    array: u16, replacement_array: u16, needles: u16, replacements: u16,
    literal: usize, stringify: usize, counted: bool,
}
fn apply_subject(chunks: &mut Vec<Chunk>, current: usize, subject: u16, search: Search) {
    let c = &mut chunks[current];
    get(c, search.array, 0);
    c.emit_if(0);
    let i = c.alloc_scratch(1);
    let n = c.alloc_scratch(1);
    let needle = c.alloc_scratch(1);
    let replacement = c.alloc_scratch(1);
    get(c, search.needles, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    call(c, "wasm:js-number", "toI32", 1, 0);
    set(c, n, 0);
    c.emit_i32_const(0, 0);
    set(c, i, 0);
    let lp = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, 0);
    let c = &mut chunks[current];
    get(c, i, 0);
    get(c, n, 0);
    c.emit_op(Op::I32_LT_U, 0);
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, 0);
    let c = &mut chunks[current];
    get(c, search.needles, 0);
    get(c, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    read_element(chunks, current, needle);
    coerce(chunks, current, needle, search.stringify);
    let c = &mut chunks[current];
    get(c, search.replacement_array, 0);
    c.emit_if(0);
    get(c, i, 0);
    get(c, search.replacements, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    call(c, "wasm:js-number", "toI32", 1, 0);
    c.emit_op(Op::I32_LT_U, 0);
    c.emit_if(0);
    get(c, search.replacements, 0);
    get(c, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    read_element(chunks, current, replacement);
    coerce(chunks, current, replacement, search.stringify);
    let c = &mut chunks[current];
    c.emit_else(0);
    c.emit_string_const("", 0);
    set(c, replacement, 0);
    c.emit_end(0);
    c.emit_else(0);
    get(c, 1, 0);
    set(c, replacement, 0);
    c.emit_end(0);
    // Both conversions run even when the needle is empty.
    for slot in [subject, needle, replacement] { get(c, slot, 0); }
    if search.counted { get(c, 3, 0); }
    direct(chunks, current, search.literal, if search.counted { 4 } else { 3 }, 0);
    let c = &mut chunks[current];
    set(c, subject, 0);
    get(c, i, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_ADD, 0);
    set(c, i, 0);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, lp, 0);
    let c = &mut chunks[current];
    c.emit_else(0);
    for slot in [subject, 0, 1] { get(c, slot, 0); }
    if search.counted { get(c, 3, 0); }
    direct(chunks, current, search.literal, if search.counted { 4 } else { 3 }, 0);
    let c = &mut chunks[current];
    set(c, subject, 0);
    c.emit_end(0);
}

fn ensure_wrapper(chunks: &mut Vec<Chunk>, counted: bool) -> usize {
    let (name, arity) = if counted { ("__stdlib_php_replace_wrapper_counted", 4) }
        else { ("__stdlib_php_replace_wrapper", 3) };
    if let Some(index) = owned(chunks, name, arity) { return index; }
    let literal = ensure_literal(chunks, counted);
    let stringify = super::string_adapter::ensure_concat_stringify_helper(chunks);
    let current = append_owned(chunks, name, arity);
    let c = &mut chunks[current];
    let search_array = c.alloc_scratch(1);
    let replacement_array = c.alloc_scratch(1);
    let subject_array = c.alloc_scratch(1);
    let needles = c.alloc_scratch(1);
    let replacements = c.alloc_scratch(1);
    let entries = c.alloc_scratch(1);
    for (input, output) in [(0, search_array), (1, replacement_array), (2, subject_array)] {
        super::copy_adapter::emit_php_arrayish_slot(c, input, 0);
        set(c, output, 0);
    }
    // Pure shallow snapshots preserve raw cells/objects. Do not pre-stringify
    // arrays: conversion may have callbacks, and repeats per subject.
    snapshot(chunks, current, 0, search_array, needles, false);
    snapshot(chunks, current, 1, replacement_array, replacements, false);
    snapshot(chunks, current, 2, subject_array, entries, true);
    for (slot, flag) in [(0, search_array), (1, replacement_array), (2, subject_array)] {
        let c = &mut chunks[current];
        get(c, flag, 0);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_if(0);
        coerce(chunks, current, slot, stringify);
        chunks[current].emit_end(0);
    }
    let c = &mut chunks[current];
    get(c, search_array, 0);
    c.emit_op(Op::I32_EQZ, 0);
    get(c, replacement_array, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_if(0);
    super::type_guard::emit_throw_const(chunks, current, "TypeError",
        "str_replace(): Argument #2 ($replace) must be of type string when argument #1 ($search) is a string", 0);
    let c = &mut chunks[current];
    c.emit_end(0);
    let search = Search { array: search_array, replacement_array, needles,
        replacements, literal, stringify, counted };
    get(c, subject_array, 0);
    c.emit_if(0);
    let output = c.alloc_scratch(1);
    let i = c.alloc_scratch(1);
    let n = c.alloc_scratch(1);
    let entry = c.alloc_scratch(1);
    let subject = c.alloc_scratch(1);
    get(c, 2, 0);
    call(c, "ecma:array", "isArray", 1, 0);
    call(c, "wasm:js-boolean", "cast", 1, 0);
    c.emit_if_value(0);
    c.emit_array_new_fixed(0, 0, 0);
    c.emit_else(0);
    call(c, "ecma:map", "new", 0, 0);
    c.emit_end(0);
    set(c, output, 0);
    get(c, entries, 0);
    c.emit_op(Op::ARRAY_LENGTH, 0);
    call(c, "wasm:js-number", "toI32", 1, 0);
    set(c, n, 0);
    c.emit_i32_const(0, 0);
    set(c, i, 0);
    let lp = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, 0);
    let c = &mut chunks[current];
    get(c, i, 0);
    get(c, n, 0);
    c.emit_op(Op::I32_LT_U, 0);
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, 0);
    let c = &mut chunks[current];
    get(c, entries, 0);
    get(c, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_TEE, entry, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    read_element(chunks, current, subject);
    coerce(chunks, current, subject, stringify);
    apply_subject(chunks, current, subject, search);
    let c = &mut chunks[current];
    get(c, output, 0);
    get(c, entry, 0);
    c.emit_i32_const(0, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    get(c, subject, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    get(c, i, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_ADD, 0);
    set(c, i, 0);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, lp, 0);
    let c = &mut chunks[current];
    get(c, output, 0);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    apply_subject(chunks, current, 2, search);
    let c = &mut chunks[current];
    get(c, 2, 0);
    c.emit_op(Op::RETURN, 0);
    current
}

pub(super) fn emit(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    assert!(matches!(argc, 3 | 4));
    let counted = argc == 4;
    let wrapper = ensure_wrapper(chunks, counted);
    let c = &mut chunks[current];
    let base = c.alloc_scratch(u16::from(argc));
    for arg in (0..u16::from(argc)).rev() { set(c, base + arg, line); }
    let accumulator = counted.then(|| c.alloc_scratch(1));
    if let Some(accumulator) = accumulator {
        c.emit_f64_const(0.0, line);
        c.emit_array_new_fixed(0, 1, line);
        set(c, accumulator, line);
    }
    for arg in 0..3 { get(c, base + arg, line); }
    if let Some(accumulator) = accumulator { get(c, accumulator, line); }
    direct(chunks, current, wrapper, argc, line);
    if let Some(accumulator) = accumulator {
        let c = &mut chunks[current];
        let result = c.alloc_scratch(1);
        let count = c.alloc_scratch(1);
        set(c, result, line);
        get(c, accumulator, line);
        c.emit_i32_const(0, line);
        c.emit_op(Op::ARRAY_GET, line);
        set(c, count, line);
        super::string_adapter::emit_str_replace_count_store(chunks, current, base + 3, count, line);
        get(&mut chunks[current], result, line);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_compiler::primitives::{bundle, globals, platforms};
    use vybe_runtime::value::ObjectKind;
    use vybe_runtime::{Value, VM};

    #[derive(Clone, Copy)]
    enum Input { Text(&'static str), Null, Bool, Number, Array, Object }
    fn physical(input: [Input; 3]) -> Result<Vec<Value>, String> {
        let mut chunks = vec![Chunk::new("replace-physical-gate")];
        let target = ensure_literal(&mut chunks, true);
        let c = &mut chunks[0];
        let acc = c.alloc_scratch(1);
        c.emit_f64_const(0.0, 0);
        c.emit_array_new_fixed(0, 1, 0);
        set(c, acc, 0);
        for input in input {
            match input {
                Input::Text(s) => c.emit_string_const(s, 0),
                Input::Null => c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
                Input::Bool => c.emit_bool_const(false, 0),
                Input::Number => c.emit_f64_const(12.0, 0),
                Input::Array => c.emit_array_new_fixed(0, 0, 0),
                Input::Object => vybe_compiler::primitives::class_slots::emit_class_alloc(c, 0),
            }
        }
        get(c, acc, 0);
        direct(&mut chunks, 0, target, 4, 0);
        let c = &mut chunks[0];
        get(c, acc, 0);
        c.emit_i32_const(0, 0);
        c.emit_op(Op::ARRAY_GET, 0);
        c.emit_array_new_fixed(0, 2, 0);
        c.emit_op(Op::RETURN, 0);
        bundle::finalize_with_runtime_helpers(&mut chunks);
        globals::declare_free_globals(&mut chunks);
        globals::normalize_global_table(&mut chunks);
        let mut vm = VM::new();
        platforms::register_platforms_all(&mut vm);
        vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
        let result = vm.run(chunks).map_err(|e| e.to_string())?;
        let Value::Object(object) = result else { panic!("expected result pair") };
        let object = object.lock().unwrap();
        let ObjectKind::Array(values) = &object.kind else { panic!("expected result array") };
        Ok(values.clone())
    }
    #[test]
    fn private_physical_string_gate_rejects_raw_values_even_for_empty_search() {
        for bad in [Input::Null, Input::Bool, Input::Number, Input::Array, Input::Object] {
            for slot in 0..3 {
                let mut input = [Input::Text("abc"), Input::Text(""), Input::Text("x")];
                input[slot] = bad;
                assert!(physical(input).is_err(), "raw argument {slot} must not be coerced");
            }
        }
    }
    #[test]
    fn private_delta_and_literal_split_join_counts() {
        for (s, n, r, out, count) in [
            ("aaaaa", "aa", "x", "xxa", 2.0),
            ("ababa", "a", "x", "xbxbx", 3.0),
            ("aa", "a", "$&", "$&$&", 2.0),
            ("abc", "", "x", "abc", 0.0),
            ("abc", "z", "", "abc", 0.0),
            ("abc", "z", "$&", "abc", 0.0),
            ("aba", "b", "XYZ", "aXYZa", 1.0),
            ("\u{1f600}a\u{1f600}", "\u{1f600}", "x", "xax", 2.0),
        ] {
            let values = physical([Input::Text(s), Input::Text(n), Input::Text(r)]).unwrap();
            assert!(matches!(&values[0], Value::String(actual) if actual.as_ref() == out));
            assert_eq!(values[1].as_f64(), count);
        }
    }
    #[test]
    fn helper_identity_requires_compiler_receipt_not_source_name() {
        let mut chunks = vec![Chunk::new("__stdlib_php_literal_replace_counted")];
        chunks[0].arity = 4;
        let target = ensure_literal(&mut chunks, true);
        assert_ne!(target, 0);
        assert_eq!(ensure_literal(&mut chunks, true), target);
        assert!(!chunks[target].takes_receiver);
        assert_eq!(chunks[target].capture_count, 0);
    }
}
