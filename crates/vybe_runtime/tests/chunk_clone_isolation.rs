use vybe_runtime::{Chunk, Value};

#[test]
fn cloned_constant_tables_keep_independent_pool_indices() {
    let mut original = Chunk::new("original");
    assert_eq!(original.intern_string_constant("shared"), 0);
    let mut clone = original.clone();
    assert_eq!(clone.intern_string_constant("shared"), 0);
    assert_eq!(clone.intern_string_constant("clone-only"), 1);
    assert_eq!(original.intern_string_constant("original-only"), 1);
    assert!(matches!(&clone.constants[1], Value::String(s) if s.as_ref()=="clone-only"));
    assert!(matches!(&original.constants[1], Value::String(s) if s.as_ref()=="original-only"));
    assert_eq!(original.intern_string_constant("clone-only"), 2);
    assert_eq!(clone.intern_string_constant("original-only"), 2);
}

#[test]
fn cloned_import_tables_rebuild_without_changing_the_original() {
    let mut original = Chunk::new("original");
    assert_eq!(original.add_import("module", "first"), 0);
    assert_eq!(original.add_import("module", "second"), 1);
    let mut clone = original.clone();
    clone.imports.remove(0);
    assert_eq!(clone.add_import("module", "second"), 0);
    assert_eq!(original.add_import("module", "second"), 1);
    assert_eq!(clone.add_import("other", "third"), 1);
    assert_eq!(original.add_import("other", "third"), 2);
    assert_eq!(original.add_global_import("globals", "first"), 0);
    let mut global_clone = original.clone();
    assert_eq!(global_clone.add_global_import("globals", "first"), 0);
    assert_eq!(global_clone.add_global_import("globals", "clone-only"), 1);
    assert_eq!(original.add_global_import("globals", "original-only"), 1);
    assert_eq!(original.add_global_import("globals", "clone-only"), 2);
    assert_eq!(global_clone.add_global_import("globals", "original-only"), 2);
}
