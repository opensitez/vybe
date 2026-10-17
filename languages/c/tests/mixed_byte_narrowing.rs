use std::sync::atomic::Ordering;
use vybe_ast::ExprKind;
use vybe_runtime::{DebugCommand, DebugRequest};

fn run_counted(chunks: Vec<vybe_runtime::Chunk>) -> u64 {
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let (commands, rx) = std::sync::mpsc::channel();
    let (events, _rx) = std::sync::mpsc::channel();
    vm.attach_debugger(rx, events, false);
    let (reply, _response) = std::sync::mpsc::channel();
    commands
        .send(DebugRequest {
            command: DebugCommand::Continue,
            reply,
        })
        .unwrap();
    let report = vm.debug_report().unwrap();
    vm.run(chunks).unwrap();
    assert_eq!(vm.pending_exit_code, 0);
    let instructions = report
        .lock()
        .unwrap()
        .live
        .instructions
        .load(Ordering::Relaxed);
    instructions
}

#[test]
fn mixed_byte_reads_narrow_only_the_unproven_branch() {
    vybe_language_c::register();
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/c/doom_startup/mixed_byte_narrowing.c");
    let bundle = vybe_compiler::projects::load(&path).unwrap();
    let optimized = bundle.prepared_module().unwrap();
    let mut baseline = optimized.clone();
    let mut restored = 0;
    for statement in &mut baseline.body {
        statement.walk_exprs_mut(&mut |value| {
            let ExprKind::Ternary { cond, then, else_ } = &value.kind else {
                return;
            };
            if !matches!(&then.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_ptr_i32_load8_u"))
            {
                return;
            }
            let ExprKind::Cast { expr, type_name } = &else_.kind else {
                return;
            };
            if type_name != "uint8" {
                return;
            }
            value.kind = ExprKind::Cast {
                expr: Box::new(vybe_ast::Expression::new(ExprKind::Ternary {
                    cond: cond.clone(),
                    then: then.clone(),
                    else_: expr.clone(),
                })),
                type_name: "uint8".into(),
            };
            restored += 1;
        });
    }
    assert!(restored > 0, "numeric branch narrowing was not removed");
    let compile = |module: vybe_ast::Module| {
        let chunks = bundle
            .compile_prepared_module(&module, &Default::default())
            .unwrap()
            .chunks;
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
        chunks
    };
    let baseline = run_counted(compile(baseline));
    let optimized = run_counted(compile(optimized));
    eprintln!("mixed byte read instructions: baseline={baseline}, optimized={optimized}");
    assert!(
        optimized + 64 < baseline,
        "snapshot rounding must not explain the saving"
    );
}
