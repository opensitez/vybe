use std::{collections::HashMap, path::Path};
use vybe_ast::{ExprKind, Expression, LambdaBody, Module, Param, StmtKind};

fn provider(module: &Module) -> (String, Vec<Param>, Expression) {
    let providers: Vec<_> = module
        .body
        .iter()
        .filter_map(|statement| {
            let StmtKind::FunctionDecl {
                name, params, body, ..
            } = &statement.kind
            else {
                return None;
            };
            if !name.starts_with("<vybe:c.charptr-visible:") {
                return None;
            }
            let [statement] = body.as_slice() else {
                panic!("provider body")
            };
            let StmtKind::Return(Some(result)) = &statement.kind else {
                panic!("provider result")
            };
            Some((name.clone(), params.clone(), result.clone()))
        })
        .collect();
    assert_eq!(
        providers.len(),
        1,
        "one producer-owned implementation per unit"
    );
    providers.into_iter().next().unwrap()
}

#[test]
fn repeated_char_pointer_formatting_shares_body_and_keeps_argument_effects() {
    vybe_language_c::register();
    let path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/shared_char_pointer_helper.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let module = bundle.prepared_module().unwrap();
    let (name, params, result) = provider(&module);
    assert_eq!(
        name, "<vybe:c.charptr-visible:0>",
        "provider identity cannot be spelled by a C source binding"
    );
    let mut baseline = module.clone();
    baseline.body.retain(|statement| {
        !matches!(&statement.kind,
        StmtKind::FunctionDecl { name: candidate, .. } if candidate == &name)
    });
    let mut calls = 0;
    for statement in &mut baseline.body {
        statement.walk_exprs_mut(&mut |expression| {
            let ExprKind::Call { callee, .. } = &mut expression.kind else {
                return;
            };
            if matches!(&callee.kind, ExprKind::Ident(candidate) if candidate == &name) {
                calls += 1;
                callee.kind = ExprKind::Lambda {
                    params: params.clone(),
                    body: LambdaBody::Expr(Box::new(result.clone())),
                    captures: vec![],
                    is_async: false,
                };
            }
        });
    }
    assert_eq!(calls, 3);
    let modules = HashMap::new();
    let before = bundle
        .compile_prepared_module(&baseline, &modules)
        .unwrap()
        .chunks;
    let after = bundle
        .compile_prepared_module(&module, &modules)
        .unwrap()
        .chunks;
    assert_eq!(after.iter().filter(|chunk| chunk.name == name).count(), 1);
    let bytes =
        |chunks: &[vybe_runtime::Chunk]| chunks.iter().map(|chunk| chunk.code.len()).sum::<usize>();
    assert!(
        bytes(&after) < bytes(&before),
        "IR bytes {} -> {}",
        bytes(&before),
        bytes(&after)
    );
    eprintln!(
        "shared char pointer IR bytes: {} -> {}",
        bytes(&before),
        bytes(&after)
    );
    for chunks in [before, after] {
        let wasm = vybe_platform_wasm::write_wasm_with_options(
            &chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&wasm)
            .unwrap();
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(
            &mut vm,
            &vybe_runtime::capabilities::Capabilities::all(),
        );
        vm.run(chunks).unwrap();
        assert_eq!(
            vm.pending_exit_code, 0,
            "each source argument must run exactly once"
        );
    }
}

#[test]
fn translation_units_without_char_pointer_formatting_emit_no_provider() {
    let module = vybe_language_c::parse("int main(void) { return 0; }").unwrap();
    assert!(
        !module.body.iter().any(|statement| matches!(&statement.kind,
        StmtKind::FunctionDecl { name, .. } if name.starts_with("<vybe:c.charptr-visible:")))
    );
}

#[test]
fn formatted_data_arguments_do_not_eagerly_read_string_pointers() {
    let module = vybe_language_c::parse(
        "#include <stdio.h>\nint main(void) { char *p = \"abc\"; char out[8]; printf(\"%.2s\", p); snprintf(out, sizeof(out), \"%.2s\", p); return 0; }",
    ).unwrap();
    assert!(!module.body.iter().any(|statement| matches!(&statement.kind,
        StmtKind::FunctionDecl { name, .. } if name.starts_with("<vybe:c.charptr-visible:"))),
        "data pointer visibility belongs to the precision-aware libc conversion");
    let mut eager = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Lambda { params, .. } = &value.kind {
                eager += usize::from(params.first().is_some_and(|param|
                    param.name == "__c_charptr_string_value"));
            }
        });
    }
    assert_eq!(eager, 0, "no inline eager visibility fallback either");
}

#[test]
fn source_dependency_shadowing_keeps_the_original_inline_scope() {
    let module = vybe_language_c::parse(
        "#include <stdio.h>\nint main(void) { int __libc_char_to_str = 1; char *p = \"abc\"; char out[4]; snprintf(out, sizeof(out), p); return 0; }",
    ).unwrap();
    assert!(
        !module.body.iter().any(|statement| matches!(&statement.kind,
        StmtKind::FunctionDecl { name, .. } if name.starts_with("<vybe:c.charptr-visible:")))
    );
    let mut inline = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Lambda { params, .. } = &value.kind {
                inline += usize::from(
                    params
                        .first()
                        .is_some_and(|p| p.name == "__c_charptr_string_value"),
                );
            }
        });
    }
    assert!(inline > 0, "keep the original scope-sensitive helper path");
}
