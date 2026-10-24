use std::path::Path;
use vybe_ast::{BindingPattern, ExprKind, StmtKind, UnaryOp};

#[test]
fn scalar_pointer_assignment_retains_pointer_value_contract() {
    let module =
        vybe_language_c::parse_with_peg("int a=1,b=2; int *p=&a; void f(void){p=&b;}").unwrap();
    let mut stores = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Assign { target, value } = &value.kind
                && matches!(&target.kind, ExprKind::Ident(name) if name == "p")
            {
                let ExprKind::Cast { type_name, expr } = &value.kind else {
                    panic!("address assignment must retain pointer-value barrier");
                };
                assert!(type_name.contains('*'));
                assert!(
                    matches!(&expr.kind, ExprKind::Unary { op: UnaryOp::AddrOf, expr }
                    if matches!(&expr.kind, ExprKind::Ident(name) if name == "b"))
                );
                stores += 1;
            }
        });
    }
    assert_eq!(stores, 1);
}

#[test]
fn scalar_character_reference_is_loaded_once_in_its_invocation() {
    let module =
        vybe_language_c::parse_with_peg("char c='q'; char *p=&c; int f(void){return *p;}").unwrap();
    let mut local = 0;
    let mut captures = 0;
    let mut loads = 0;
    let mut signed_bytes = 0;
    for statement in &module.body {
        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            assert!(!declarations.iter().any(
                |declaration| matches!(&declaration.pattern, BindingPattern::Ident(name)
                    if name.starts_with("__c_char_deref_pointer"))
            ));
        }
        if let StmtKind::FunctionDecl { name, body, .. } = &statement.kind
            && name == "f"
        {
            for statement in body {
                if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
                    local += declarations
                        .iter()
                        .filter(|declaration| {
                            matches!(&declaration.pattern, BindingPattern::Ident(name)
                            if name.starts_with("__c_char_deref_pointer"))
                        })
                        .count();
                }
                statement.walk_exprs(&mut |value| match &value.kind {
                    ExprKind::Binary {
                        op: vybe_ast::BinOp::Integer(vybe_ast::IntOp::ShrS, vybe_ast::BitLane::W32),
                        left, right,
                    } if matches!(&right.kind, ExprKind::Lit(vybe_ast::Literal::Int(24))) => {
                        assert!(matches!(&left.kind, ExprKind::Binary {
                            op: vybe_ast::BinOp::Integer(vybe_ast::IntOp::Shl, vybe_ast::BitLane::W32),
                            right, ..
                        } if matches!(&right.kind, ExprKind::Lit(vybe_ast::Literal::Int(24)))));
                        signed_bytes += 1;
                    }
                    ExprKind::Assign { target, value }
                        if matches!(&target.kind, ExprKind::Ident(name)
                                if name.starts_with("__c_char_deref_pointer")) =>
                    {
                        assert!(matches!(&value.kind, ExprKind::Cast { type_name, expr }
                                if type_name == "void*"
                                    && matches!(&expr.kind, ExprKind::Ident(name) if name == "p")));
                        captures += 1;
                    }
                    ExprKind::RefLoad(pointer)
                        if matches!(&pointer.kind, ExprKind::Ident(name)
                                if name.starts_with("__c_char_deref_pointer")) =>
                    {
                        loads += 1
                    }
                    _ => {}
                });
            }
        }
    }
    assert_eq!((local, captures, loads), (1, 1, 1));
    assert_eq!(signed_bytes, 1);
}

#[test]
fn scalar_pointer_carriers_match_native_peg_and_static_wasm() {
    vybe_language_c::register();
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tests/c");
    for name in [
        "pointers_basic/pointer_can_reference_char_value.c",
        "pointers_basic/pointer_can_be_reassigned_to_other_variable.c",
        "doom_startup/scalar_pointer_carriers.c",
    ] {
        let source = root.join(name);
        let native = std::env::temp_dir().join(format!(
            "vybe-scalar-carrier-{}-{}",
            std::process::id(),
            source.file_stem().unwrap().to_str().unwrap()
        ));
        let compiled = std::process::Command::new("cc")
            .args(["-std=c11", "-O0"])
            .arg(&source)
            .arg("-o")
            .arg(&native)
            .output()
            .unwrap();
        assert!(
            compiled.status.success(),
            "{name}: {}",
            String::from_utf8_lossy(&compiled.stderr)
        );
        let executed = std::process::Command::new(&native).output().unwrap();
        assert!(
            executed.status.success(),
            "{name}: {}",
            String::from_utf8_lossy(&executed.stderr)
        );
        let parsers: [(&str, vybe_runtime::registry::ParseFn); 3] = [
            ("peg", vybe_language_c::parse_with_peg),
            ("native", vybe_language_c::parse_optimized),
            ("pratt", vybe_language_c::parse_contextual_pratt),
        ];
        let mut reference_wasm = None;
        for (backend, parser) in parsers {
            let mut bundle = vybe_compiler::projects::load(&source).unwrap();
            bundle.language.parse = parser;
            let chunks = bundle.compile().unwrap();
            for chunk in &chunks {
                assert!(!chunk.name.contains("dyneq"));
                assert!(
                    !vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
                    "{name} {backend}: {}",
                    chunk.name
                );
            }
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: false,
                },
            );
vybe_platform_wasm::validate_binary(&bytes).unwrap();
            let module = vybe_platform_wasm::inspection::CoreModule::new(&bytes).unwrap();
            let mut abstract_imports = Vec::new();
            module.imports(|module, name, ty| {
                if matches!(ty, vybe_platform_wasm::inspection::ImportType::Func { .. }) {
                    abstract_imports.push(module == "ecma:value"
                        && matches!(name, "abstractEq" | "abstractNe"));
                }
                Ok(())
            }).unwrap();
            module.function_bodies(|body| {
                for instruction in body.instructions()? {
                    if let Some(function_index) = instruction?.function_index() {
                        assert_ne!(abstract_imports.get(function_index as usize), Some(&true),
                            "emitted WASM references abstract equality at function {function_index}");
                    }
                }
                Ok(())
            }).unwrap();
            if let Some(reference) = &reference_wasm {
                assert!(
                    reference == &bytes,
                    "{name}: {backend} standard WASM differs from PEG"
                );
            } else {
                reference_wasm = Some(bytes);
            }
            let mut vm = vybe_runtime::VM::new();
            vybe_runtime::init_all_registered(
                &mut vm,
                &vybe_runtime::capabilities::Capabilities::all(),
            );
            vm.run(chunks).unwrap();
            assert_eq!(vm.pending_exit_code, 0, "{name}: {backend}");
        }
    }
}
