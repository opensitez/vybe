use std::sync::{Arc, Mutex};
use vybe_runtime::{
    Chunk, Op, VM, Value,
    value::{Function, Object, ObjectKind},
};

fn input(kind: &str) -> Value {
    let mut object = Object::new();
    object.kind = match kind {
        "object" => ObjectKind::Ordinary,
        "array" => ObjectKind::Array(vec![]),
        "function" => ObjectKind::Function(Function {
            name: None,
            arity: 0,
            chunk_index: 0,
            upvalues: vec![],
        }),
        "host" => ObjectKind::HostFunction(0),
        "iterator" => ObjectKind::StringViewIter {
            source: Arc::from("é😀"),
            byte_offset: 0,
            scalar_offset: 0,
        },
        "null" => return Value::Null,
        "number" => return Value::I32(42),
        "string" => return Value::String(Arc::from("text")),
        _ => panic!("unknown input kind"),
    };
    Value::Object(Arc::new(Mutex::new(object)))
}

fn worker(repeats: usize) -> (VM, Value) {
    let mut root = Chunk::new("shape-guard-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut body = Chunk::new("shape-guard");
    body.arity = 1;
    body.param_count = 1;
    body.local_count = 2;
    body.emit_i32_const(0, 0);
    body.emit_op_u16(Op::LOCAL_SET, 1, 0);
    for _ in 0..repeats {
        body.emit_op_u16(Op::LOCAL_GET, 1, 0);
        body.emit_op_u16(Op::LOCAL_GET, 0, 0);
        body.emit_ref_type_op(
            Op::REF_TEST,
            vybe_runtime::opcode::heaptype::HeapType::Abstract(
                vybe_runtime::opcode::heaptype::HT_STRUCT,
            ),
            0,
        );
        body.emit_op(Op::I32_ADD, 0);
        body.emit_op_u16(Op::LOCAL_SET, 1, 0);
    }
    body.emit_op_u16(Op::LOCAL_GET, 1, 0);
    body.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    let function = vm.run(vec![root, body]).unwrap();
    (vm, function)
}

#[test]
fn struct_guard_distinguishes_objects_arrays_and_functions() {
    let (mut vm, function) = worker(7);
    for kind in [
        "object", "array", "function", "host", "iterator", "null", "number", "string",
    ] {
        let expected = if matches!(kind, "object" | "iterator") {
            7
        } else {
            0
        };
        assert_eq!(
            vm.invoke(&function, &[input(kind)]).unwrap(),
            Value::I32(expected),
            "{kind}"
        );
    }
    for value in [
        Value::Undefined,
        Value::Bool(false),
        Value::I64(1),
        Value::F32(1.0),
        Value::F64(1.0),
        Value::TypedNull(1),
        Value::V128([0; 16]),
    ] {
        assert_eq!(vm.invoke(&function, &[value]).unwrap(), Value::I32(0));
    }
}

#[test]
#[ignore = "loaded debug VM shape-guard measurement; compare frozen/current binaries"]
fn loaded_struct_guard_sample() {
    let kind = std::env::var("VYBE_STRUCT_KIND").unwrap();
    let (mut vm, function) = worker(64);
    let args = [input(&kind)];
    let expected = Value::I32(if matches!(kind.as_str(), "object" | "iterator") {
        64
    } else {
        0
    });
    for _ in 0..32 {
        assert_eq!(vm.invoke(&function, &args).unwrap(), expected);
    }
    let start = std::time::Instant::now();
    for _ in 0..2000 {
        assert_eq!(
            std::hint::black_box(vm.invoke(&function, &args).unwrap()),
            expected
        );
    }
    println!(
        "STRUCT_GUARD kind={kind} probes=128000 elapsed_ms={}",
        start.elapsed().as_secs_f64() * 1000.0
    );
}
