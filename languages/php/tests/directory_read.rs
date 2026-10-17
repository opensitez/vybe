//! Directory branches produce exactly one value and leave no receiver behind.
use std::sync::{Arc, Mutex};
use vybe_runtime::{Chunk, Op, VM, Value};
use vybe_runtime::value::Object;

pub fn directory_chunks() -> Vec<Chunk> {
    let mut root = Chunk::new("directory-owner");
    // Use the same primitive comparison helper and module-global binding as
    // production compilation; the standalone adapter consumes that binding.
    vybe_compiler::primitives::functions::emit_ref_func(&mut root, 2, 0, 0);
    vybe_compiler::primitives::globals::emit_write(&mut root,"__vybe_dynlt",0);
    vybe_compiler::primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut read = Chunk::new("directory-read");
    read.arity = 1; read.param_count = 1; read.local_count = 1;
    read.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut chunks = vec![root, read];
    vybe_language_php::emitter::filesystem_adapter::emit_dir_read(&mut chunks, 1, 1, 0);
    chunks[1].emit_op(Op::RETURN, 0);
    chunks.push(vybe_compiler::primitives::ops::build_dyn_comparison("__stdlib_dynlt"));
    vybe_compiler::primitives::globals::declare_free_globals(&mut chunks);
    vybe_compiler::primitives::globals::normalize_global_table(&mut chunks);
    chunks
}

#[test]
fn directory_entries_exhaustion_and_index() {
    vybe_language_php::register();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let read = vm.run(directory_chunks()).unwrap();
    for names in [vec![], vec!["a"], vec!["a", "b", "c"]] {
        let entries = Value::Object(Arc::new(Mutex::new(Object::new_array(names.iter().map(|s|Value::String((*s).into())).collect()))));
        let mut object = Object::new();
        object.set("__entries".into(), entries);
        object.set("__index".into(), Value::F64(0.0));
        let directory = Value::Object(Arc::new(Mutex::new(object)));
        for (index, name) in names.iter().enumerate() {
            let value = vm.invoke(&read, &[directory.clone()]).unwrap();
            assert!(matches!(value, Value::String(ref s) if s.as_ref() == *name));
            let Value::Object(object) = &directory else { unreachable!() };
            assert_eq!(object.lock().unwrap().get("__index").as_f64(), (index+1) as f64);
        }
        for _ in 0..3 {
            assert!(matches!(vm.invoke(&read, &[directory.clone()]).unwrap(),Value::Bool(false)));
        }
        let Value::Object(object) = directory else { unreachable!() };
        assert_eq!(object.lock().unwrap().get("__index").as_f64(), names.len() as f64);
    }
}
