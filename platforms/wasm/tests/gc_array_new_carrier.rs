use vybe_runtime::{Chunk, Op, Value, VM};

fn program(kind: usize, count: i32, checked: bool, index: Option<i32>) -> Vec<Chunk> {
    let mut c = Chunk::new("array-new-carrier");
    match kind {
        0 => c.emit_i32_const(0, 7),
        1 => c.emit_f64_const(-0.0, 7),
        2 => c.emit_bool_const(false, 7),
        3 => c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 7),
        _ => unreachable!(),
    }
    c.emit_i32_const(count, 7);
    if checked {
        let f = c.add_import("wasm:js-number", "toI32");
        c.emit_call(f, 1, 7);
    }
    c.emit_op_u16(Op::ARRAY_NEW, 0, 7);
    if let Some(index) = index {
        c.emit_i32_const(index, 7);
        c.emit_op(Op::ARRAY_GET, 7);
    } else {
        c.emit_op(Op::ARRAY_LENGTH, 7);
    }
    c.emit_op(Op::RETURN, 7);
    vec![c]
}

#[test]
fn vm_positive_counts_and_exact_initializer_tags() {
    for kind in 0..4 {
        for count in [0, 1, 128] {
            for checked in [false, true] {
                let mut v = VM::new();
                vybe_runtime::js_builtins::register(&mut v);
                assert_eq!(v.run(program(kind, count, checked, None)).unwrap(), Value::I32(count));
                if count > 0 {
                    for index in [0, count - 1] {
                        let actual = v.run(program(kind, count, checked, Some(index))).unwrap();
                        match (kind, actual) {
                            (0, Value::I32(0)) | (2, Value::Bool(false)) | (3, Value::Null) => (),
                            (1, Value::F64(n)) => assert_eq!(n.to_bits(), (-0.0f64).to_bits()),
                            (kind, actual) => panic!("initializer kind {kind}: {actual:?}"),
                        }
                    }
                }
            }
        }
    }
}

#[test]
fn generic_array_new_boxed_count_restores_standard_i32_boundary() {
 for kind in 0..4 {for count in [0,1,128] {for checked in [false,true] {
  let mut indices=vec![None];if count>0 {indices.extend([Some(0),Some(count-1)]);}
  for index in indices {
   let bytes=vybe_platform_wasm::write_wasm(&program(kind,count,checked,index));
   wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
    .validate_all(&bytes).unwrap_or_else(|e|panic!("kind{kind} count{count} checked{checked} index{index:?}: {e}"));
  }
 }}}
}
