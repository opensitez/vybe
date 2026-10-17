use std::{collections::HashMap, path::Path, sync::{Arc, Mutex}};
use vybe_ast::{BindingPattern, ClassMember, ClassModifiers, ExprKind, Expression, Statement, StmtKind, StorageAccess, TypeHint};
const CLASS: &str = "__c_private_process_state";
fn main() {
    vybe_language_c::register();
    let source = std::env::args().nth(1).expect("source path");
    let bundle = vybe_compiler::projects::load(Path::new(&source)).unwrap();
    let baseline = bundle.prepared_module().unwrap();
    let mut candidate = baseline.clone();
    let fields = ["exit_status", "exit_callbacks", "quick_callbacks", "shutdown_started", "quick_shutdown", "skip_atexit", "skip_flush", "stdout_buffer"];
    let mut replaced = 0;
    for statement in &mut candidate.body {
        if let StmtKind::VarDecl { declarations, .. } = &mut statement.kind {
            for declaration in declarations {
                if matches!(&declaration.pattern, BindingPattern::Ident(name) if name == "__c_process") {
                    assert!(matches!(declaration.init.as_ref().map(|value| &value.kind), Some(ExprKind::Object(_))));
                    let mut hint = TypeHint::checked(CLASS);
                    hint.access = Some(StorageAccess::Ordinary);
                    declaration.type_hint = Some(hint);
                    declaration.init = Some(Expression::new(ExprKind::New {
                        class: Box::new(Expression::new(ExprKind::Ident(CLASS.into()))), args: vec![],
                    }));
                    replaced += 1;
                }
            }
        }
    }
    assert_eq!(replaced, 1);
    candidate.body.insert(0, Statement::new(StmtKind::ClassDecl {
        name: CLASS.into(), parents: vec![], interfaces: vec![], decorators: vec![],
        modifiers: ClassModifiers { is_sealed: true, ..Default::default() },
        members: fields.into_iter().map(|name| ClassMember::Field {
            name: name.into(), type_hint: None, init: None, modifiers: Default::default(),
            with_events: false, array_bounds: None, storage: None,
        }).collect(),
    }));
    let mut reference = None;
    for (label, module) in [("object", baseline), ("class", candidate)] {
        let compiled = bundle.compile_prepared_module(&module, &HashMap::new()).unwrap();
        let bytecode: usize = compiled.chunks.iter().map(|chunk| chunk.code.len()).sum();
        let bytes = vybe_platform_wasm::write_wasm(&compiled.chunks);
        std::fs::write(format!("/private/tmp/vega-c-process-{label}.wasm"), &bytes).unwrap();
        println!("{label}: bytecode={bytecode} wasm={}", bytes.len());
        for chunk in &compiled.chunks {
            if chunk.name.starts_with("__c_") { println!("  {}: {}", chunk.name, chunk.code.len()); }
        }
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
        let output = Arc::new(Mutex::new(Vec::new()));
        let captured = output.clone();
        vm.register_host_fn("wasi:cli/stdout", "write-via-stream", Box::new(move |ctx, args| {
            captured.lock().unwrap().extend(ctx.stream_drain(&args[0]));
            let (future, id) = ctx.create_future();
            ctx.resolve_future(id, vybe_runtime::Value::Null);
            future
        }));
        vm.run(compiled.chunks).unwrap();
        let result = (vm.pending_exit_code, output.lock().unwrap().clone());
        println!("  result={result:?}");
        if let Some(expected) = &reference { assert_eq!(&result, expected); }
        else { reference = Some(result); }
    }
}
