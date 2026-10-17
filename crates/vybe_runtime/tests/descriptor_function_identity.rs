use std::sync::Arc;
use vybe_runtime::chunk::{ConstExpr, TypeEntry};
use vybe_runtime::{Chunk, Op, Value, VM};

#[test]
fn constant_function_references_share_identity_across_vtable_initializers() {
    let mut entry = Chunk::new("descriptor-function-identity");
    entry.types = vec![
        TypeEntry { name: "DescriptorIdentityBase".into(), ..Default::default() },
        TypeEntry { name: "DescriptorIdentityDerived".into(), parent_index: 1, ..Default::default() },
        TypeEntry {
            name: "#desc DescriptorIdentityBase".into(),
            describes_index: 1,
            fields: vec!["__desc_proto".into(), "__desc_props".into(), "method".into()],
            ..Default::default()
        },
        TypeEntry {
            name: "#desc DescriptorIdentityDerived".into(),
            describes_index: 2,
            parent_index: 3,
            fields: vec!["__desc_proto".into(), "__desc_props".into(), "method".into()],
            ..Default::default()
        },
    ];
    for (name, typeidx) in [("descriptor_base_method", 3), ("descriptor_derived_method", 4)] {
        entry.add_global_init(name, ConstExpr::StructNew {
            typeidx,
            fields: Arc::from([
                ConstExpr::Value(Value::Null),
                ConstExpr::Value(Value::Null),
                ConstExpr::RefFunc(1),
            ]),
        });
    }
    entry.emit_i32_const(0, 1);
    entry.emit_op(Op::RETURN, 1);

    let mut method = Chunk::new("shared-method");
    method.emit_i32_const(7, 1);
    method.emit_op(Op::RETURN, 1);

    let mut vm = VM::new();
    vm.run(vec![entry, method]).unwrap();
    let Some(Value::Object(base)) = vm.global("descriptor_base_method") else {
        panic!("base initializer did not produce a descriptor");
    };
    let Some(Value::Object(derived)) = vm.global("descriptor_derived_method") else {
        panic!("derived initializer did not produce a descriptor");
    };
    assert!(!Arc::ptr_eq(base, derived), "different classes need distinct descriptors");
    let base_method = base.lock().unwrap().fields[2].clone();
    let derived_method = derived.lock().unwrap().fields[2].clone();
    let (Value::Object(base_method), Value::Object(derived_method)) = (base_method, derived_method) else {
        panic!("descriptor method slots did not contain function references");
    };
    assert!(Arc::ptr_eq(&base_method, &derived_method));
}
