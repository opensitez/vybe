use vybe_compiler::primitives::globals::{declare_free_globals, emit_read, emit_write,
    finalize_global_table, normalize_global_table};
use vybe_runtime::{Chunk, Value};
use vybe_runtime::chunk::{ConstExpr, GlobalInit};

fn fixture() -> Vec<Chunk> {
    let mut entry = Chunk::new("entry");
    emit_read(&mut entry, "local_callable", 1);
    emit_read(&mut entry, "local_number", 1);
    emit_read(&mut entry, "external_value", 1);
    emit_write(&mut entry, "written_value", 1);
    entry.global_inits.push(GlobalInit { name: "local_callable".into(), init: ConstExpr::RefFunc(1) });
    let mut function = Chunk::new("function");
    function.global_inits.push(GlobalInit { name: "local_number".into(), init: ConstExpr::Value(Value::I32(7)) });
    function.global_inits.push(GlobalInit { name: "metadata_only".into(), init: ConstExpr::RefFunc(1) });
    vec![entry, function]
}

#[test]
fn module_initializers_are_definitions_in_separate_and_fused_finalization() {
    let mut separate = fixture();
    declare_free_globals(&mut separate);
    assert_eq!(separate[0].global_imports.len(), 1);
    assert_eq!(separate[0].global_imports[0].name, "external_value");
    normalize_global_table(&mut separate);
    let mut fused = fixture();
    finalize_global_table(&mut fused);
    for (left, right) in separate.iter().zip(&fused) {
        assert_eq!(left.code, right.code);
        assert_eq!(left.global_imports, right.global_imports);
        assert_eq!(left.globals, right.globals);
    }
    assert_eq!(fused[0].globals.iter().filter(|name| name.as_str() == "metadata_only").count(), 1);
    assert!(std::sync::Arc::ptr_eq(&fused[0].globals, &fused[1].globals));
}

#[test]
fn explicit_imports_are_retained_despite_same_named_initializers() {
    for fused in [false, true] {
        let mut chunks = fixture();
        chunks[1].add_global_import("explicit:provider", "local_callable");
        chunks[0].add_global_import("env", "local_number");
        let explicit = chunks.iter().map(|chunk| chunk.global_imports.clone()).collect::<Vec<_>>();
        if fused { finalize_global_table(&mut chunks); }
        else { declare_free_globals(&mut chunks); normalize_global_table(&mut chunks); }
        for (chunk, original) in chunks.iter().zip(explicit) {
            for import in original { assert!(chunk.global_imports.contains(&import)); }
        }
        assert!(!chunks[0].global_imports.iter().any(|import|
            import.module == "env" && import.name == "local_callable"));
        for name in ["local_callable", "local_number"] {
            assert_eq!(chunks[0].globals.iter().filter(|entry| entry.as_str() == name).count(), 1);
        }
    }
}

#[test]
fn initializer_only_modules_get_shared_global_indices_without_code() {
    let mut chunks = vec![Chunk::new("entry"), Chunk::new("function")];
    for chunk in &mut chunks {
        chunk.global_inits.push(GlobalInit { name: "only_metadata".into(), init: ConstExpr::RefFunc(1) });
    }
    finalize_global_table(&mut chunks);
    assert!(chunks.iter().all(|chunk| chunk.code.is_empty() && chunk.global_imports.is_empty()));
    assert_eq!(chunks[0].globals.iter().filter(|name| name.as_str() == "only_metadata").count(), 1);
    assert!(std::sync::Arc::ptr_eq(&chunks[0].globals, &chunks[1].globals));
}
