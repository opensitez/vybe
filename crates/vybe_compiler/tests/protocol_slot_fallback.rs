use vybe_compiler::primitives::{bundle, class_slots, expressions, functions, globals, platforms};
use vybe_runtime::{Chunk, Op, VMError, Value, VM};

#[derive(Clone, Copy)]
enum SlotValue {
    Absent,
    Null,
    Zero,
    False,
    Method,
}

fn subject(chunk: &mut Chunk, key: &str, value: SlotValue) {
    class_slots::emit_class_alloc(chunk, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    match value {
        SlotValue::Absent => return,
        SlotValue::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        SlotValue::Zero => chunk.emit_i32_const(0, 0),
        SlotValue::False => chunk.emit_bool_const(false, 0),
        SlotValue::Method => functions::emit_ref_func(chunk, 1, 0, 0),
    }
    let key = class_slots::resolve_interned(
        chunk,
        &class_slots::ClassSlot::internal(key),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        chunk,
        class_slots::ObjSource::Local(0),
        &key,
        class_slots::ValueSource::Stack,
        0,
    );
}

fn run(
    emit: impl FnOnce(&mut Vec<Chunk>),
    method_arity: u8,
    method_result: Value,
) -> Result<Value, VMError> {
    let mut entry = Chunk::new("<script>");
    entry.local_count = 2;
    let mut method = functions::create_function_chunk("protocol_method", method_arity);
    method.local_count = u16::from(method_arity);
    match method_result {
        Value::String(value) => method.emit_string_const(&value, 0),
        Value::F64(value) => method.emit_f64_const(value, 0),
        Value::I32(value) => method.emit_i32_const(value, 0),
        _ => panic!("unsupported test result"),
    }
    method.emit_op(Op::RETURN, 0);
    let mut chunks = vec![entry, method];
    emit(&mut chunks);
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    let bytes = vybe_platform_wasm::write_wasm_with_options(
        &chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_vybe_metadata: false,
            include_name_section: true,
        },
    );
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .unwrap();
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm.run(chunks)
}

fn stringify(value: SlotValue) -> Result<Value, VMError> {
    run(
        |chunks| {
            subject(
                &mut chunks[0],
                &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
                value,
            );
            expressions::emit_rich_to_string(&mut chunks[0], 0, 0);
        },
        1,
        Value::String("custom".into()),
    )
}

#[test]
fn stringification_missing_slots_fall_back_and_methods_remain_callable() {
    for value in [SlotValue::Absent, SlotValue::Null] {
        assert_eq!(
            stringify(value).unwrap(),
            Value::String("[object Object]".into())
        );
    }
    assert_eq!(
        stringify(SlotValue::Method).unwrap(),
        Value::String("custom".into())
    );
    for value in [SlotValue::Zero, SlotValue::False] {
        assert!(
            stringify(value).is_err(),
            "present noncallable slot must not fall back"
        );
    }
}

#[test]
fn primitive_stringification_falls_back_without_operator_calls() {
    let value = run(
        |chunks| {
            chunks[0].emit_f64_const(13.0, 0);
            chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
            expressions::emit_rich_to_string(&mut chunks[0], 0, 0);
        },
        1,
        Value::String("unused".into()),
    )
    .unwrap();
    assert_eq!(value, Value::String("13".into()));
}

fn greater(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_GT, line);
}

#[test]
fn primitive_comparison_missing_slots_reaches_numeric_fallback() {
    let value = run(
        |chunks| {
            chunks[0].emit_f64_const(14.0, 0);
            chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
            chunks[0].emit_f64_const(0.0, 0);
            chunks[0].emit_op_u16(Op::LOCAL_SET, 1, 0);
            expressions::emit_rich_compare_locals(
                chunks,
                0,
                0,
                1,
                &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Gt),
                expressions::RichFallback::Op(greater),
                0,
                false,
            );
        },
        2,
        Value::F64(1.0),
    )
    .unwrap();
    assert_eq!(value.as_i32(), 1);
}

#[test]
fn known_numeric_comparison_bypasses_protocol_lookup() {
    let value = run(
        |chunks| {
            chunks[0].emit_f64_const(14.0, 0);
            chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
            chunks[0].emit_f64_const(0.0, 0);
            chunks[0].emit_op_u16(Op::LOCAL_SET, 1, 0);
            let start = chunks[0].code.len();
            expressions::emit_rich_compare_locals(
                chunks,
                0,
                0,
                1,
                &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Gt),
                expressions::RichFallback::Op(greater),
                0,
                true,
            );
            let mut expected = Chunk::new("expected");
            expected.emit_op_u16(Op::LOCAL_GET, 0, 0);
            expected.emit_op_u16(Op::LOCAL_GET, 1, 0);
            expected.emit_op(Op::F64_GT, 0);
            assert_eq!(&chunks[0].code[start..], expected.code.as_slice());
        },
        2,
        Value::F64(1.0),
    )
    .unwrap();
    assert_eq!(value.as_i32(), 1);
}

#[test]
fn comparison_keeps_declared_operator_and_compare_methods() {
    for key in [
        vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Gt),
        "CompareTo".into(),
    ] {
        let value = run(
            |chunks| {
                subject(&mut chunks[0], &key, SlotValue::Method);
                chunks[0].emit_f64_const(0.0, 0);
                chunks[0].emit_op_u16(Op::LOCAL_SET, 1, 0);
                expressions::emit_rich_compare_locals(
                    chunks,
                    0,
                    0,
                    1,
                    &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Gt),
                    expressions::RichFallback::Op(greater),
                    0,
                    false,
                );
            },
            2,
            Value::F64(1.0),
        )
        .unwrap();
        assert_eq!(value.as_i32(), 1);
    }
}

fn arithmetic_fallback(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_f64_const(5.0, line);
}

fn unary_fallback(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::DROP, line);
    chunk.emit_f64_const(5.0, line);
}

#[test]
fn arithmetic_and_unary_slots_preserve_fallback_and_present_methods() {
    for value in [SlotValue::Absent, SlotValue::Null, SlotValue::Method] {
        let expected = if matches!(value, SlotValue::Method) {
            7
        } else {
            5
        };
        let arithmetic = run(
            |chunks| {
                subject(&mut chunks[0], "__add__", value);
                chunks[0].emit_f64_const(1.0, 0);
                chunks[0].emit_op_u16(Op::LOCAL_SET, 1, 0);
                expressions::emit_rich_arithmetic(
                    &mut chunks[0],
                    0,
                    1,
                    "__add__",
                    arithmetic_fallback,
                    0,
                );
            },
            2,
            Value::F64(7.0),
        )
        .unwrap();
        assert_eq!(arithmetic.as_i32(), expected);
        let unary = run(
            |chunks| {
                subject(&mut chunks[0], "__neg__", value);
                expressions::emit_rich_unary(&mut chunks[0], 0, "__neg__", unary_fallback, 0);
            },
            1,
            Value::F64(7.0),
        )
        .unwrap();
        assert_eq!(unary.as_i32(), expected);
    }
}

#[test]
fn bool_and_length_slots_handle_absence_and_preserve_methods() {
    for value in [SlotValue::Absent, SlotValue::Null, SlotValue::Method] {
        let boolean = run(
            |chunks| {
                subject(&mut chunks[0], "__bool__", value);
                expressions::emit_rich_bool(&mut chunks[0], 0, 0);
            },
            1,
            Value::I32(0),
        )
        .unwrap();
        assert_eq!(
            boolean.as_i32(),
            if matches!(value, SlotValue::Method) {
                0
            } else {
                1
            }
        );
    }
    let length = run(
        |chunks| {
            chunks[0].emit_string_const("abc", 0);
            chunks[0].emit_op_u16(Op::LOCAL_SET, 0, 0);
            expressions::emit_smart_length(&mut chunks[0], 0, 0);
        },
        1,
        Value::I32(4),
    )
    .unwrap();
    assert_eq!(length.as_i32(), 3);
    let custom_length = run(
        |chunks| {
            subject(
                &mut chunks[0],
                &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Len),
                SlotValue::Method,
            );
            expressions::emit_smart_length(&mut chunks[0], 0, 0);
        },
        1,
        Value::I32(4),
    )
    .unwrap();
    assert_eq!(custom_length.as_i32(), 4);
}
