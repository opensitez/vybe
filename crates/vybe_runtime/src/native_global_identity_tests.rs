use crate::chunk::{ConstExpr, imported_global_key};
use crate::{Chunk, Op, VM, Value};
use std::sync::Arc;

fn native_module() -> Vec<Chunk> {
    let imported = imported_global_key("env", "seed");
    let mut script = Chunk::new("native-module");
    script.native_global_import_count = Some(1);
    script.global_imports.push(crate::chunk::Import {
        module: "env".into(),
        name: "seed".into(),
    });
    script.globals = Arc::new(vec![imported.clone(), "local".into(), "derived".into()].into());
    script.add_global_init("local", ConstExpr::GlobalGet(imported));
    script.add_global_init(
        "derived",
        ConstExpr::Add(
            Box::new(ConstExpr::GlobalGet("local".into())),
            Box::new(ConstExpr::Value(Value::I32(5))),
        ),
    );
    // Mutation must remain inside this instance, and not replace the host binding.
    script.emit_i32_const(99, 0);
    script.emit_op_u32(Op::GLOBAL_SET, 1, 0);
    script.emit_op_u32(Op::GLOBAL_GET, 2, 0);
    script.emit_op(Op::RETURN, 0);
    vec![script]
}

#[test]
fn native_defined_globals_are_fresh_on_both_loader_paths() {
    for loader in 0..3 {
        let mut vm = VM::new();
        let imported = imported_global_key("env", "seed");
        vm.set_global(&imported, Value::I32(37));
        // A source module's intentionally shared binding must not be replaced.
        vm.set_global("local", Value::I32(17));
        for instance in 0..2 {
            let result = match loader {
                0 => vm.run(native_module()),
                1 => vm.run_linked(native_module(), vec![]),
                _ => {
                    let prepared = crate::PreparedLinkedCode::prepare(native_module(), usize::MAX)
                        .ok()
                        .unwrap();
                    vm.run_linked_nested_prepared(prepared, vec![])
                }
            };
            assert_eq!(result.unwrap(), Value::I32(42));
            assert_eq!(
                vm.global(&format!("__wasm_instance_{instance}_global_1")),
                Some(&Value::I32(99))
            );
            assert_eq!(
                vm.global(&format!("__wasm_instance_{instance}_global_2")),
                Some(&Value::I32(42))
            );
            assert_eq!(vm.global(&imported), Some(&Value::I32(37)));
            assert_eq!(vm.global("local"), Some(&Value::I32(17)));
        }
    }
}

#[test]
fn source_chunks_keep_existing_shared_global_policy() {
    let mut vm = VM::new();
    for initial in [17, 91] {
        let mut script = Chunk::new("source-module");
        script.globals = Arc::new(vec!["shared".into()].into());
        script.add_global_init("shared", ConstExpr::Value(Value::I32(initial)));
        script.emit_op_u32(Op::GLOBAL_GET, 0, 0);
        script.emit_op(Op::RETURN, 0);
        assert_eq!(vm.run(vec![script]).unwrap(), Value::I32(17));
    }
}
