// Private native gate compiled inside the normal Compiler library.
// The external driver registers the real PHP consumer against this same library.
use super::{RuntimeCompilerService,VM,Value};
use std::sync::Arc;
fn configured_vm()->VM { let mut vm=VM::new(); crate::primitives::platforms::register_platforms_all(&mut vm); vm }
fn assert_numeric_value(value:Value,expected:f64) {match value { Value::I32(v)=>assert_eq!(v as f64,expected), Value::I64(v)=>assert_eq!(v as f64,expected), Value::F64(v)=>assert_eq!(v,expected), other=>panic!("expected number, got {other:?}") }}
fn v6_permit(_:crate::languages::Language)->bool {true}

fn v6_completed_pair(runtime:&mut super::PhpIncludeRuntime,vm:&VM,
    language:crate::languages::Language,pool:&crate::preparation::PreparationPool,
    path:&std::path::Path,reject_certificate:bool,
)->(Arc<std::sync::atomic::AtomicBool>,Arc<super::AtomicUsize>) {
    let registration=crate::frontend_preparation::snapshot(language.parse);
    let context=Arc::new(super::FusedLoweringContext::capture(vm,language,v6_permit,None,registration).unwrap());
    runtime.completion_permission=Some((super::FrontendIdentity::of(language),v6_permit));
    runtime.completion_permission_override=true;runtime.select_completion_permission(language);
    runtime.fused_context=Some(context.clone());
    let budget=Arc::new(super::AtomicUsize::new(0));let work_budget=budget.clone();
    let cancelled=Arc::new(std::sync::atomic::AtomicBool::new(false));let work_cancelled=cancelled.clone();
    let (send,result)=crossbeam_channel::bounded(1);
    let post:crate::preparation::PostPrepare=Box::new(move|prepared| {
        if work_cancelled.load(std::sync::atomic::Ordering::Relaxed){return;}
        let lowered=context.lower_prepared(prepared,language,&work_budget)
            .expect("actual PHP must be admitted to worker lowering");
        let pending=lowered.compilation.as_ref().expect("actual worker compilation must succeed");
        let dependencies=pending.compilation.compile_dependencies.as_ref().expect("worker certificate");
        if reject_certificate {dependencies.certificate.lock().unwrap().reject();}
        if !work_cancelled.load(std::sync::atomic::Ordering::Relaxed){send.send(lowered).unwrap();}
    });
    let ticket=pool.submit_with_post_prepare(path.to_path_buf(),language,None,Some(post)).unwrap();
    let deadline=std::time::Instant::now()+std::time::Duration::from_secs(30);
    while !ticket.is_ready(){assert!(std::time::Instant::now()<deadline,"paired post-lowering must complete");std::thread::yield_now();}
    assert!(!result.is_empty(),"positive actual emission, not a merely parsed ticket");
    assert!(budget.load(std::sync::atomic::Ordering::Relaxed)>0,"fixture must own real prepared link metadata before consume/cancel");
    runtime.preparing.insert(path.to_path_buf(),ticket);
    runtime.lowering.insert(path.to_path_buf(),super::LoweringTicket{result,cancelled:cancelled.clone(),link_budget:budget.clone()});
    (cancelled,budget)
}

fn v6_paired_completion_has_positive_hit_and_source_certificate_fallback() {
    let language=crate::languages::find_by_name("php").unwrap();
    let base=std::env::temp_dir().join(format!("vybe-v6-paired-{}",uuid::Uuid::new_v4()));
    std::fs::create_dir_all(&base).unwrap();let path=base.join("unit.php");
    let pool=crate::preparation::PreparationPool::new(1,4).unwrap();
    let mut vm=configured_vm();super::ensure_php_runtime_registered(&mut vm);
    let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
    for (edit,reject,expected) in [(false,false,7.0),(true,false,9.0),(false,true,7.0)] {
        std::fs::write(&path,"<?php return 7;").unwrap();let path=path.canonicalize().unwrap();
        let before=runtime.completion_handoff_hits;
        let (cancelled,budget)=v6_completed_pair(&mut runtime,&vm,language,&pool,&path,reject);
        if edit {std::fs::write(&path,"<?php return 9;").unwrap();}
        let source=std::fs::read_to_string(&path).unwrap();let bundle=super::bundle_from_source(source,language,&path);
        let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&path).unwrap();
        assert_eq!(runtime.completion_handoff_hits-before,usize::from(!edit&&!reject),"actual validated reuse versus fallback");
        assert!(cancelled.load(std::sync::atomic::Ordering::Relaxed));
        assert_eq!(budget.load(std::sync::atomic::Ordering::Relaxed),0,"metadata owner released on consume/reject");
        assert_numeric_value(RuntimeCompilerService::new(&mut vm).run_compiled(compiled).unwrap(),expected);
    }
    std::fs::remove_dir_all(base).unwrap();
}

fn v6_cancelled_paired_receiver_releases_metadata_and_preserves_ready_ast() {
    let language=crate::languages::find_by_name("php").unwrap();
    let path=std::env::temp_dir().join(format!("vybe-v6-cancel-{}.php",uuid::Uuid::new_v4()));
    std::fs::write(&path,"<?php return 41;").unwrap();let path=path.canonicalize().unwrap();
    let pool=crate::preparation::PreparationPool::new(1,1).unwrap();
    let mut vm=configured_vm();super::ensure_php_runtime_registered(&mut vm);
    let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
    let (cancelled,budget)=v6_completed_pair(&mut runtime,&vm,language,&pool,&path,false);
    drop(runtime.lowering.remove(&path).unwrap());
    assert!(cancelled.load(std::sync::atomic::Ordering::Relaxed));assert_eq!(budget.load(std::sync::atomic::Ordering::Relaxed),0);
    assert!(runtime.preparing[&path].is_ready());
    let bundle=super::bundle_from_source(std::fs::read_to_string(&path).unwrap(),language,&path);
    let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&path).unwrap();
    assert_eq!(runtime.completion_handoff_hits,0);
    assert_numeric_value(RuntimeCompilerService::new(&mut vm).run_compiled(compiled).unwrap(),41.0);
    std::fs::remove_file(path).unwrap();
}

fn v6_foreground_committed_receipt_ranks_only_opted_in_service_and_all_load_forms_keep_scope() {
    let language=crate::languages::find_by_name("php").unwrap();
    let base=std::env::temp_dir().join(format!("vybe-v6-current-{}",uuid::Uuid::new_v4()));
    std::fs::create_dir_all(&base).unwrap();
    std::fs::write(base.join("child.php"),"<?php $x += 1;").unwrap();
    let main=base.join("main.php");
    let source="<?php $x=0; include __DIR__.'/child.php'; require __DIR__.'/child.php'; include_once __DIR__.'/child.php'; require_once __DIR__.'/child.php'; return $x;";
    std::fs::write(&main,source).unwrap();
    let mut vm=configured_vm();
    let mut service=RuntimeCompilerService::new(&mut vm);
    service.set_completion_lowering(language,Some(v6_permit)).unwrap();
    let compiled=service.compile_source_by_name(source,"php",&main).unwrap();
    assert!(service.php_runtime.current_source_rank_receipts>0,"actual foreground source ranking");
    assert!(service.php_runtime.committed_foreground_receipts>0,"real registered PHP preparation receipt, not generic lookahead");
    assert!(service.php_runtime.committed_foreground_ready_paths>0,"committed receipt must resolve real include targets, not an empty protocol batch");
    assert_numeric_value(service.run_compiled(compiled).unwrap(),2.0);
    service.set_completion_lowering(language,None).unwrap();
    let before=service.php_runtime.current_source_rank_receipts;
    let compiled=service.compile_source_by_name("<?php return 3;","php",base.join("disabled.php")).unwrap();
    assert_eq!(service.php_runtime.current_source_rank_receipts,before,"explicit disable must bypass ranking");
    assert_numeric_value(service.run_compiled(compiled).unwrap(),3.0);
    drop(service);std::fs::remove_dir_all(base).unwrap();
}

pub fn run() {
 let language=crate::languages::find_by_name("php").expect("external real PHP consumer must register first");
 let selected=crate::frontend_preparation::snapshot(language.parse);
 assert!(selected.capabilities.prepare.is_some(),"PHP must publish preparation to this same normal Compiler instance");
 assert!(!selected.capabilities.early_hints,"gate must exercise committed receipt, not legacy early scout");
 v6_paired_completion_has_positive_hit_and_source_certificate_fallback();
 v6_cancelled_paired_receiver_releases_metadata_and_preserves_ready_ast();
 v6_foreground_committed_receipt_ranks_only_opted_in_service_and_all_load_forms_keep_scope();
 println!("V6 same-Compiler real PHP completion/certificate/source/cancel/foreground receipts passed");
}
