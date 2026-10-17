use std::collections::HashMap;
use std::path::Path;
use vybe_ast::{ExprKind, Expression, LambdaBody, Module, Param, StmtKind};

fn restore_immediate_lambdas(module: &mut Module) {
    let mut helpers: HashMap<String, (Vec<Param>, Expression)> = HashMap::new();
    module.body.retain(|statement| {
        let StmtKind::FunctionDecl {
            name, params, body, ..
        } = &statement.kind
        else {
            return true;
        };
        if !name.starts_with("__c_static_iife_") {
            return true;
        }
        let [result] = body.as_slice() else {
            panic!("unexpected static helper body")
        };
        let StmtKind::Return(Some(result)) = &result.kind else {
            panic!("unexpected static helper result")
        };
        helpers.insert(name.clone(), (params.clone(), result.clone()));
        false
    });
    fn expand(value: &mut Expression, helpers: &HashMap<String, (Vec<Param>, Expression)>) {
        let ExprKind::Call { callee, args, .. } = &mut value.kind else {
            return;
        };
        let ExprKind::Ident(name) = &callee.kind else {
            return;
        };
        if name == "__c_invoke_fixed" {
            assert!(!args.is_empty());
            *callee = Box::new(args.remove(0).value);
            expand(value, helpers);
            return;
        }
        let Some((params, result)) = helpers.get(name) else {
            return;
        };
        let mut result = result.clone();
        result.walk_exprs_mut(&mut |value| expand(value, helpers));
        callee.kind = ExprKind::Lambda {
            params: params.clone(),
            body: LambdaBody::Expr(Box::new(result)),
            is_async: false,
            captures: vec![],
        };
    }
    for statement in &mut module.body {
        statement.walk_exprs_mut(&mut |value| expand(value, &helpers));
    }
}

fn opcode_count(chunks: &[vybe_runtime::Chunk]) -> usize {
    chunks
        .iter()
        .map(|chunk| {
            vybe_runtime::debug::disassemble(chunk)
                .lines()
                .filter(|line| line.as_bytes().first().is_some_and(u8::is_ascii_digit))
                .count()
        })
        .sum()
}

fn dynamic_equality_references(chunks: &[vybe_runtime::Chunk]) -> usize {
    chunks
        .iter()
        .map(|chunk| {
            let dump = vybe_runtime::debug::disassemble(chunk);
            dump.lines()
                .filter(|line| line.contains("dyneq") || line.contains("dynamic_equal"))
                .count()
                + usize::from(chunk.name.contains("dyneq") || chunk.name.contains("dynamic_equal"))
        })
        .sum()
}

#[test]
fn static_helpers_preserve_computed_pointer_aliases() {
    vybe_language_c::register();
    for fixture in [
        "static_pointer_helper.c",
        "computed_managed_alias.c",
        "static_helpers_collision.c",
        "fixed_invocation_order.c",
    ] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures")
            .join(fixture);
        let bundle = vybe_compiler::projects::load(&path).unwrap();
        let optimized = bundle.prepared_module().unwrap();
        let mut fixed_calls = 0;
        for statement in &optimized.body {
            statement.walk_exprs(&mut |value| {
                if matches!(&value.kind, ExprKind::Call { callee, .. }
                    if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_invoke_fixed"))
                {
                    fixed_calls += 1;
                }
            });
        }
        assert!(fixed_calls > 0, "{fixture}: no fixed invocation exercised");
        let mut baseline = optimized.clone();
        restore_immediate_lambdas(&mut baseline);
        let modules = HashMap::new();
        let before = bundle
            .compile_prepared_module(&baseline, &modules)
            .unwrap()
            .chunks;
        let chunks = bundle
            .compile_prepared_module(&optimized, &modules)
            .unwrap()
            .chunks;
        let before_count = opcode_count(&before);
        let after_count = opcode_count(&chunks);
        assert!(
            after_count < before_count,
            "{before_count} -> {after_count}"
        );
        eprintln!("{fixture} total opcodes: {before_count} -> {after_count}");
        for chunks in [before, chunks] {
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: true,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap_or_else(|error| panic!("{fixture}: {error}"));
            let mut vm = vybe_runtime::VM::new();
            vybe_runtime::init_all_registered(
                &mut vm,
                &vybe_runtime::capabilities::Capabilities::all(),
            );
            vm.run(chunks).unwrap();
            assert_eq!(vm.pending_exit_code, 0);
        }
    }
}

#[test]
#[ignore = "full Doom before/after compilation; run when the shared build lane is free"]
fn static_helpers_doom_emitted_opcode_census() {
    vybe_language_c::register();
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/chocolate-doom/vybe_doom_core.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let optimized = bundle.prepared_module().unwrap();
    let mut fixed_calls = 0;
    for statement in &optimized.body {
        statement.walk_exprs(&mut |value| {
            if matches!(&value.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_invoke_fixed"))
            {
                fixed_calls += 1;
            }
        });
    }
    assert!(fixed_calls > 0, "no actual Doom helpers qualified");
    let mut baseline = optimized.clone();
    restore_immediate_lambdas(&mut baseline);
    let modules = HashMap::new();
    let before = bundle
        .compile_prepared_module(&baseline, &modules)
        .unwrap()
        .chunks;
    let after = bundle
        .compile_prepared_module(&optimized, &modules)
        .unwrap()
        .chunks;
    let before_count = opcode_count(&before);
    let after_count = opcode_count(&after);
    let before_dyneq = dynamic_equality_references(&before);
    let after_dyneq = dynamic_equality_references(&after);
    assert_eq!(before_dyneq, 0, "baseline dynamic equality references");
    assert_eq!(after_dyneq, 0, "normalized dynamic equality references");
    assert!(
        after_count < before_count,
        "{before_count} -> {after_count}"
    );
    eprintln!("Doom fixed literal calls: {fixed_calls}; total emitted opcodes: {before_count} -> {after_count}");
    eprintln!(
        "Doom chunk counts: {} -> {}; dyneq references: {before_dyneq} -> {after_dyneq}",
        before.len(),
        after.len()
    );
    for chunks in [before, after] {
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            &chunks,
            vybe_platform_wasm::WasmWriteOptions {
                include_vybe_metadata: false,
                include_name_section: true,
            },
        );
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
    }
}
