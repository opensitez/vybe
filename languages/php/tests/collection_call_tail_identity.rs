//! PHP integration gate for opt-in shared conditional-call recipes.
//! Parse once, retain the same AST in every compiler lane, validate native Wasm,
//! and compare actual raw/prepared VM outcomes including error source frames.
use vybe_ast::{ExprKind, Module};
use vybe_runtime::{Chunk, PreparedLinkedCode, VM, registry::CompilerOptimizations};

fn parse_fixture(source: &str) -> Module {
    vybe_language_php::register();
    let mut module = vybe_language_php::parse(source).unwrap();
    // The PHP unknown-method lowering naturally clones the tail with identical
    // spans. Explicit source ternaries have different spans, so manufacture that
    // exact shape once in the fixture AST, before either compilation lane.
    for statement in &mut module.body {
        statement.walk_exprs_mut(&mut |expr| {
            if let ExprKind::Ternary { then, else_, .. } = &mut expr.kind
                && matches!(&then.kind, ExprKind::Call { .. })
                && let ExprKind::Ternary { else_: tail, .. } = &mut else_.kind {
                *tail = then.clone();
            }
        });
    }
    module
}

fn compile(module: &Module, arrays: bool, interpolation: bool) -> Vec<Chunk> {
    let chunks = vybe_compiler::Compiler::with_profile(
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap(),
    ).with_optimizations(CompilerOptimizations {
        array_call_tail_identity: arrays,
        interpolation_call_tail_identity: interpolation,
        ..Default::default()
    }).with_run_module_end_hooks(false).compile(module).unwrap();
    let wasm = vybe_platform_wasm::write_wasm_with_options(&chunks,
        vybe_platform_wasm::WasmWriteOptions {include_vybe_metadata:false,include_name_section:true});
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm).unwrap();
    chunks
}

fn outcome(chunks: Vec<Chunk>, prepared: bool) -> String {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_runtime::framework::init_all_on_vm(&mut vm, &[&vybe_language_php::Plugin]);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);
    let resolved = chunks.first().into_iter().flat_map(|chunk| &chunk.imports)
        .map(|import| match vm.resolve_import_target(&import.module,&import.name) {
            Ok(target) => Ok(target),
            Err(_) if matches!(import.module.as_str(),"env"|"*") =>
                Ok(vybe_runtime::ImportTarget::StdlibRedirect(import.name.clone())),
            Err(error) => Err(error),
        }).collect::<Result<Vec<_>,_>>().unwrap();
    let result = if prepared {
        let code=PreparedLinkedCode::prepare(chunks,64*1024*1024).ok().expect("prepared admission");
        vm.run_linked_prepared(code,resolved)
    } else {vm.run_linked(chunks,resolved)};
    match result {
        Ok(value) => format!("Ok({value:?})"),
        Err(error) => {
            // Shrinking code changes byte offsets; retain raw diagnostics while
            // comparing exact message/source line and ordered frame identities.
            eprintln!("raw call-tail error prepared={prepared}: {error:?}");
            let frames=error.call_stack.iter().map(|frame|
                (&frame.chunk_name,frame.line)).collect::<Vec<_>>();
            format!("Err({:?},{:?},{frames:?})",error.message,error.line)
        },
    }
}

fn bytes(chunks: &[Chunk]) -> usize {chunks.iter().map(|c|c.code.len()).sum()}
fn check(source: &str, arrays: bool, interpolation: bool, expected: Option<&str>, reduction: bool) -> String {
    check_module(parse_fixture(source),arrays,interpolation,expected,reduction)
}
fn check_module(module: Module, arrays: bool, interpolation: bool, expected: Option<&str>, reduction: bool) -> String {
    let before_ast=format!("{module:?}");
    let old=compile(&module,false,false);
    let selected=compile(&module,arrays,interpolation);
    assert_eq!(before_ast,format!("{module:?}"),"emission must not rewrite AST");
    // Public function identity/order and calling convention remain unchanged.
    let identities=|chunks:&[Chunk]|chunks.iter().map(|c|
        (c.name.clone(),c.arity,c.result_arity,c.takes_receiver,c.capture_count,c.is_async,c.is_generator)
    ).collect::<Vec<_>>();
    assert_eq!(identities(&old),identities(&selected));
    if reduction {assert!(bytes(&selected)<bytes(&old),"fixture must exercise new admission");}
    let oracle=outcome(old.clone(),false);
    if let Some(expected)=expected {assert_eq!(oracle,expected);}
    assert_eq!(outcome(old,true),oracle);
    assert_eq!(outcome(selected.clone(),false),oracle);
    assert_eq!(outcome(selected,true),oracle);
    oracle
}

const PREFIX: &str = "<?php $events=''; function pred($tag,$value){global $events;$events.=$tag;return $value;} function arg($tag){global $events;$events.=$tag;return $tag;} function f($value){global $events;$events.='F';return $value;} function g($value){global $events;$events.='G';return $value;} ";

#[test]
fn selected_predicates_and_array_arguments_run_in_order_once() {
    for (a,b,events) in [(true,true,"AXF"),(true,false,"AXF"),(false,true,"ABYG"),(false,false,"ABXF")] {
        let source=format!("{PREFIX}$r=pred('A',{})?f([arg('X')]):(pred('B',{})?g([arg('Y')]):f([arg('X')]));return $events;",if a{"true"}else{"false"},if b{"true"}else{"false"});
        check(&source,true,false,Some(&format!("Ok(String({events:?}))")),true);
    }
}

#[test]
fn interpolation_is_independent_and_retains_selected_call_order() {
    use vybe_ast::{Expression,InterpolPart};
    let source=format!("{PREFIX}$r=pred('A',false)?f([arg('X'),arg('Y')]):(pred('B',false)?g([arg('Z'),arg('W')]):f([arg('X'),arg('Y')]));return $events;");
    let mut module=parse_fixture(&source);
    let mut converted=0;
    for statement in &mut module.body {
        statement.walk_exprs_mut(&mut |expr| {
            if let ExprKind::Array(elements)=&expr.kind && elements.len()==2 {
                // Explicit AST fixture: PHP module-scope interpolation wraps
                // variable reads in unsupported NullCoalesce. Preserve the real
                // parsed side-effecting Call operands, guarantee Interpolation,
                // and keep spans unchanged before either compiler lane sees it.
                let first=elements[0].value.clone();
                let second=elements[1].value.clone();
                *expr=Expression::with_span(ExprKind::Interpolation(vec![
                    InterpolPart::Text("prefix ".into()),InterpolPart::Expr(first),
                    InterpolPart::Formatted(second,"s".into()),
                    InterpolPart::Text(" suffix".into()),
                ]),expr.span);
                converted+=1;
            }
        });
    }
    assert_eq!(converted,3,"fixture must carry all three interpolation tails");
    check_module(module.clone(),false,true,Some("Ok(String(\"ABXYF\"))"),true);
    let old=compile(&module,false,false);let array_only=compile(&module,true,false);
    assert!(vybe_platform_wasm::write_wasm(&old)==vybe_platform_wasm::write_wasm(&array_only),"Array opt-in must not enable interpolation");
}

#[test]
fn named_spread_and_reference_arguments_preserve_baseline() {
    for source in [
        "<?php function f($value){return $value[0];} $r=true?f(value:[7]):(false?f([8]):f([7]));return $r;",
        "<?php function f($value){return $value[0];} $r=true?f(...[[7]]):(false?f([8]):f(...[[7]]));return $r;",
        "<?php function f($value){$value[0]=9;} $x=1;$r=true?f([&$x]):(false?f([2]):f([&$x]));return $x;",
    ] {check(source,true,false,None,false);}
}

#[test]
fn nullsafe_receiver_and_error_frames_preserve_baseline() {
    check("<?php $events='';function arg(){global $events;$events.='X';return 1;}$o=null;$r=true?$o?->f([arg()]):(false?0:$o?->f([arg()]));return $events;",true,false,None,false);
    let error=check("<?php function arg(){throw new Exception('call-tail-sentinel');} function f($x){return $x;}\n$r=false?f([arg()]):(false?0:f([arg()]));return $r;",true,false,None,true);
    assert!(error.starts_with("Err(") && error.contains("call-tail-sentinel"),
        "must exercise the native source throw: {error}");
    assert!(!error.contains("no active runtime"), "harness failure is not a guest error");
}

#[test]
fn default_options_keep_original_emission_recipe() {
    let module=parse_fixture("<?php function f($a){return $a;}return true?f([1]):(false?f([2]):f([1]));");
    let old=compile(&module,false,false);
    let implicit=vybe_compiler::Compiler::with_profile(vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap())
        .with_run_module_end_hooks(false).compile(&module).unwrap();
    assert!(vybe_platform_wasm::write_wasm(&old)==vybe_platform_wasm::write_wasm(&implicit));
}

#[test]
fn actual_unknown_receiver_runs_receiver_and_array_argument_once() {
    check("<?php $events='';class O {function f($x){global $events;$events.='F';return $x;}} function receiver($x){global $events;$events.='R';return $x;}function arg(){global $events;$events.='X';return 7;}$o=new O();$r=receiver($o)->f([arg()]);return $events;",true,false,Some("Ok(String(\"RXF\"))"),true);
}
