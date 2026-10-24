//! Regression for the public prepared-manifest route, not only Compiler's
//! lower-level owned-AST entry. Run with --features service-verification.
#![cfg(feature = "service-verification")]
use vybe_compiler::dynamic::FusedManifestValidation;
use vybe_compiler::primitives::deferred_class_bodies::take_statistics;
use vybe_runtime::VM;
#[path = "../examples/emission_import_interning/canonical.rs"]
mod metadata;

#[test]
fn prepared_manifest_preserves_source_and_donates_only_owned_class_bodies() {
    vybe_language_php::register();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm, &[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    let directory = std::env::temp_dir().join(format!("vybe-owned-manifest-route-{}", uuid::Uuid::new_v4()));
    std::fs::create_dir_all(&directory).unwrap();
    let path = directory.join("class.php");
    std::fs::write(&path, "<?php namespace ManifestGate; class Cell { public function read(){return 7;} } return 7;").unwrap();
    let mut expected = None;
    for owned in [true, false] {
        let mut prepared = FusedManifestValidation::prepare(&vm, &path, owned).unwrap();
        let source = prepared.source_state().unwrap();
        take_statistics();
        let unit = prepared.emit().unwrap();
        let statistics = take_statistics();
        assert_eq!(source, prepared.source_state().unwrap(), "source manifest changed");
        assert_eq!(statistics.methods > 0, owned, "owned route lost class donation or borrowed input was consumed: {statistics:?}");
        let actual = unit.chunks.iter().map(|chunk| metadata::canonical(&format!("{chunk:?}"))).collect::<Vec<_>>();
        if let Some(expected) = &expected { assert_eq!(&actual, expected, "complete emitted metadata differs"); }
        else { expected = Some(actual); }
    }
    std::fs::remove_dir_all(directory).unwrap();
}
