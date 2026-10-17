use super::*;

#[test]
fn record_callback_return_inference_uses_owner_declaration_and_preserves_pointer_depth() {
    let mut walker = Walker::default();
    walker.structs.insert("C".into(), vec!["run".into(), "plain".into(), "pointer".into(), "alias".into()]);
    walker.struct_field_types.insert("C".into(), HashMap::from([
        ("run".into(), "int*".into()), ("plain".into(), "int*".into()),
        ("pointer".into(), "int**".into()), ("alias".into(), "Callback".into()),
    ]));
    walker.struct_function_pointer_fields.insert("C".into(), HashSet::from(["run".into(), "pointer".into()]));
    walker.var_types.insert("p".into(), "struct C *".into());
    walker.typedef_function_return_types.insert("Callback".into(), "double".into());
    walker.function_return_types.insert("run".into(), "double".into());
    assert_eq!(walker.infer_function_return_type(&member(ident("p"), "run")).as_deref(), Some("int"));
    assert_eq!(walker.infer_function_return_type(&member(ident("p"), "pointer")).as_deref(), Some("int*"));
    assert_eq!(walker.infer_function_return_type(&member(ident("p"), "alias")).as_deref(), Some("double"));
    assert!(walker.infer_function_return_type(&member(ident("p"), "plain")).is_none());
    assert!(walker.declared_callback_return_type("int*", false).is_none());
    walker.function_return_types.insert("owner_once".into(), "struct C *".into());
    for source in ["p->run(side_effect())", "owner_once()->run(side_effect())"] {
        let pair = CParser::parse(Rule::postfix_expression, source).unwrap().next().unwrap();
        let call = pair.clone().into_inner().last().unwrap();
        let types = walker.unevaluated_callback_call_returns(pair);
        assert_eq!(types.len(), 1);
        assert_eq!(types.get(&call.as_span().start()).map(String::as_str), Some("int"));
    }
    assert_eq!(walker.tmp_counter, 0);
    assert!(walker.generated_local_declarations.is_empty());
}

#[test]
fn record_callback_results_execute_standard_wasm_without_dynamic_inequality() {
    crate::register();
    for source in [
        include_str!("../../../tests/c/doom_startup/frontend_call_handoffs.c"),
        "typedef unsigned short (*F)(void);struct C{F run;};unsigned short value(void){return 60001;}int main(){struct C c;c.run=value;return c.run()!=60001;}",
        "struct C{unsigned long long (*run)(void);};unsigned long long value(void){return 4294967297ULL;}int main(){struct C c;c.run=value;return c.run()!=4294967297ULL;}",
        "int run(void){return 9;}struct C{double (*run)(void);};double value(void){return 1.5;}int main(){struct C c;c.run=value;return c.run()!=1.5;}",
        "struct C{int (*run)(void);};int calls;struct C c;int value(void){return 42;}struct C *owner_once(void){++calls;return &c;}int main(){c.run=value;return owner_once()->run()!=42 || calls!=1;}",
    ] {
        let expected = format!("{:?}", crate::parse_with_peg(source).unwrap());
        let mut baseline = None;
        for parse in [crate::parse_with_peg, crate::parse_optimized, crate::parse_owned_ast,
            crate::parse_owned_contextual, crate::parse_contextual_pratt]
        {
            let module = parse(source).unwrap();
            assert_eq!(format!("{module:?}"), expected, "{source}");
            let profile = vybe_compiler::profile::parse_profile(crate::profile_source()).unwrap();
            let chunks = vybe_compiler::primitives::Compiler::with_profile(profile).compile(&module).unwrap();
            for chunk in &chunks {
                let dump = vybe_runtime::debug::disassemble(chunk);
                assert!(!dump.contains("__vybe_dyneq"), "{}: {source}\n{dump}", chunk.name);
            }
            let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
                vybe_platform_wasm::WasmWriteOptions { include_vybe_metadata: false, include_name_section: true });
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes).unwrap();
            if let Some(expected) = &baseline { assert_eq!(&bytes, expected, "{source}"); }
            else { baseline = Some(bytes); }
            let mut vm = vybe_runtime::VM::new();
            vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
            vm.run(chunks).unwrap();
            assert_eq!(vm.pending_exit_code, 0, "{source}");
        }
    }
}
