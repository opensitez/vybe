use super::parse;
use vybe_ast::canon::ComponentBinaryDef as D;

fn binary(source: &str) -> vybe_ast::canon::ComponentBinary {
    let module = parse(source).unwrap_or_else(|error| panic!("{source}: {error}"));
    assert!(module.canon.component_source);
    assert!(
        module.body.is_empty(),
        "worker definitions must not become runtime shims"
    );
    let binary = module.canon.binary.expect("ordered worker definitions");
    let bytes = vybe_platform_wasm::write_component_binary(&binary).unwrap();
    vybe_platform_wasm::validate_binary(&bytes).unwrap();
    assert_eq!(
        *binary,
        vybe_platform_wasm::read_component_binary(&bytes).unwrap()
    );
    (*binary).clone()
}

#[test]
fn worker_availability_preserves_share_flag_and_core_function_binder() {
    for (flag, shared) in [("", false), ("shared", true)] {
        let component = binary(&format!(
            "(component
            (canon thread.available-parallelism {flag} (core func $available))
            (core instance $exports (export \"available\" (func $available))))"
        ));
        assert!(
            matches!(component.definitions[0], D::CanonThreadAvailableParallelism { shared: actual } if actual == shared)
        );
        let D::CoreInstanceFromExports(exports) = &component.definitions[1] else {
            panic!("core exports")
        };
        assert_eq!(exports[0].index, 0);
    }
}

#[test]
fn worker_spawn_ref_preserves_type_width_and_named_indices() {
    for (flag, shared) in [("", false), ("shared", true)] {
        for value in ["i32", "i64"] {
            let function = if shared {
                format!("(shared (func (param {value})))")
            } else {
                format!("(func (param {value}))")
            };
            let component = binary(&format!(
                "(component
                (core type $entry {function})
                (canon thread.spawn-ref {flag} $entry (core func $spawn)))"
            ));
            assert!(
                component
                    .definitions
                    .iter()
                    .any(|definition| matches!(definition,
                D::CanonThreadSpawnRef { shared: actual, type_index: 0 } if *actual == shared))
            );
        }
    }
}

#[test]
fn worker_spawn_indirect_preserves_named_table_and_type_indices() {
    let component = binary(
        "(component
        (core type $entry (shared (func (param i32))))
        (core module $tables (table (export \"workers\") shared 0 1 (ref null (shared func))))
        (core instance $tables-instance (instantiate $tables))
        (alias core export $tables-instance \"workers\" (core table $workers))
        (canon thread.spawn-indirect shared $entry $workers (core func $spawn)))",
    );
    assert!(component.definitions.iter().any(|definition| matches!(
        definition,
        D::CanonThreadSpawnIndirect {
            shared: true,
            type_index: 0,
            table_index: 0
        }
    )));
}

#[test]
fn worker_inline_and_nested_forms_use_resolved_primary_type_spaces() {
    let component = binary(
        "(component
        (core type $entry (func (param i64)))
        (core func $spawn (canon thread.spawn-ref $entry))
        (canon thread.available-parallelism (core func $available))
        (core instance (export \"spawn\" (func $spawn))
                       (export \"available\" (func $available))))",
    );
    let D::CoreInstanceFromExports(exports) = component.definitions.last().unwrap() else {
        panic!()
    };
    assert_eq!(
        exports
            .iter()
            .map(|export| export.index)
            .collect::<Vec<_>>(),
        [0, 1]
    );
    let nested = binary(
        "(component
        (core type $entry (func (param i32)))
        (component $child
            (alias outer 1 $entry (core type $copy))
            (canon thread.spawn-ref $copy (core func $spawn))))",
    );
    let D::NestedComponent(child) = &nested.definitions[1] else {
        panic!()
    };
    assert!(child.definitions.iter().any(|definition| matches!(
        definition,
        D::CanonThreadSpawnRef {
            shared: false,
            type_index: 0
        }
    )));
}

#[test]
fn invalid_worker_source_does_not_fall_back_to_executable_shims() {
    for ty in [
        "(func)",
        "(func (param f64))",
        "(func (param i32) (result i32))",
        "(shared (func (param i32)))",
    ] {
        assert!(
            parse(&format!(
                "(component (core type $entry {ty})
            (canon thread.spawn-ref $entry (core func $spawn)))"
            ))
            .is_err()
        );
    }
}
