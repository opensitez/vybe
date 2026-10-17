const SOURCE: &str = include_str!("../../../tests/c/doom_startup/immutable_declared_functions.c");

#[test]
fn c_declared_functions_state_immutable_binding_policy() {
    for parse in [crate::parse_with_peg, crate::parse_optimized] {
        assert_eq!(parse(SOURCE).unwrap().directives.declared_function_bindings_live,
            Some(false));
    }
}

#[test]
fn completed_c_declaration_uses_direct_edge_without_freezing_callbacks() {
    crate::register();
    for parse in [crate::parse_with_peg, crate::parse_optimized] {
        let module = parse(SOURCE).unwrap();
        let profile = vybe_compiler::profile::parse_profile(crate::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile_with_imports(&module).unwrap().chunks;
        let target = chunks.iter().position(|chunk| chunk.name == "target").unwrap();
        let caller = chunks.iter().find(|chunk| chunk.name == "direct_caller").unwrap();
        let mut live_module = module.clone();
        live_module.directives.declared_function_bindings_live = Some(true);
        let live_profile = vybe_compiler::profile::parse_profile(crate::profile_source()).unwrap();
        let live_chunks = vybe_compiler::primitives::Compiler::with_profile(live_profile)
            .compile_with_imports(&live_module).unwrap().chunks;
        let live_caller = live_chunks.iter().find(|chunk| chunk.name == "direct_caller").unwrap();
        assert!(live_caller.direct_defined_calls.values().all(|edge| !edge.source));
        eprintln!("C caller bytecode bytes: live={} immutable={}",
            live_caller.code.len(), caller.code.len());
        assert!(caller.direct_defined_calls.values().any(|edge|
            edge.source && edge.target as usize == target),
            "completed declaration must have a real direct call: {}",
            vybe_runtime::debug::disassemble(caller));
        assert_eq!(chunks[target].direct_entry_has_captures, Some(false));
        for name in ["shadowed", "indirect_caller"] {
            let chunk = chunks.iter().find(|chunk| chunk.name == name).unwrap();
            assert!(chunk.direct_defined_calls.values().all(|edge|
                edge.target as usize != target), "mutable callback frozen in {name}");
        }
    }
}

#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
fn immutable_c_declarations_execute_identical_validated_standard_wasm() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized], [SOURCE]);
}
