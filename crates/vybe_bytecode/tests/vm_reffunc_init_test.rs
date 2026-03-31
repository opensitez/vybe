use vybe_bytecode::*;
use vybe_bytecode::chunk::*;
use vybe_bytecode::value::ObjectKind;
use std::rc::Rc;

#[test]
fn global_init_reffunc_creates_function() {
    let mut vm = VM::new();
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    script.global_inits.push(GlobalInit {
        name: "__test_fn".to_string(),
        init: ConstExpr::RefFunc(1),
    });
    let name_c = script.add_constant(Value::String(Rc::from("__test_fn")));
    script.emit_op_u16(opcode::Op::global_get, name_c, 0);
    script.emit_op(opcode::Op::halt, 0);

    let mut func_chunk = Chunk::new("test_func");
    func_chunk.arity = 0;
    func_chunk.local_count = 1;
    let c42 = func_chunk.add_constant(Value::I32(42));
    func_chunk.emit_op_u16(opcode::Op::r#const, c42, 0);
    func_chunk.emit_op(opcode::Op::r#return, 0);

    let result = vm.run(vec![script, func_chunk]).unwrap();
    match &result {
        Value::Object(obj) => {
            let o = obj.borrow();
            assert!(matches!(&o.kind, ObjectKind::Function(_)), "should be a Function, got {:?}", o.kind);
        }
        other => panic!("expected Function object, got {:?}", other),
    }
}
