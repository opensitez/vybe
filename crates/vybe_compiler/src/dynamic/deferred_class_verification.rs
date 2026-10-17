//! Feature-only service checks. The PHP integration target supplies a VM with
//! the actual PHP plugin initialized; Compiler never depends on that frontend.
use super::*;
fn numeric(value: Value, expected: i32) {
    let actual=match value {Value::I32(n)=>n as f64,Value::I64(n)=>n as f64,Value::F64(n)=>n,other=>panic!("expected number, got {other:?}")};
    assert_eq!(actual,expected as f64);
}
fn deny_completion(_:Language)->bool {false}
fn allow_completion(_:Language)->bool {true}
struct Scratch(std::path::PathBuf);
impl Drop for Scratch {fn drop(&mut self){let _=std::fs::remove_dir_all(&self.0);}}
fn write(base:&Scratch,name:&str,source:&str)->PathBuf {
    let path=base.0.join(name);std::fs::write(&path,source).unwrap();path.canonicalize().unwrap()
}
fn prepared(pool:&crate::preparation::PreparationPool,path:&Path,language:Language)->crate::preparation::PreparedFile {
    pool.submit(path.to_path_buf(),language).unwrap().wait().unwrap()
}

pub(super) fn verify(mut vm:VM) {
    let language=crate::languages::find_by_name("php").unwrap();
    let registration=crate::frontend_preparation::snapshot(language.parse);
    assert!(registration.capabilities.deferred_class_bodies.is_some());
    let base=Scratch(std::env::temp_dir().join(format!("vybe-owned-class-{}",uuid::Uuid::new_v4())));
    std::fs::create_dir(&base.0).unwrap();ensure_php_runtime_registered(&mut vm);
    let mut runtime=PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
    let _active_runtime=runtime.activate(&mut vm,None);
    runtime.completion_permission=Some((FrontendIdentity::of(language),deny_completion));
    runtime.completion_permission_override=true;

    // Force committed foreground compilation rather than a fused borrowed
    // completion. The ordinary method, constructor and caller scope execute.
    let source="<?php class OwnedDonationServiceA{public function value(){return 41;}}return (new OwnedDonationServiceA)->value();";
    let path=write(&base,"foreground.php",source);let bundle=bundle_from_source(source,language,path.clone());
    crate::primitives::deferred_class_bodies::take_statistics();
    let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&path).unwrap();
    let stats=crate::primitives::deferred_class_bodies::take_statistics();
    assert!(stats.methods>0,"actual committed include must donate: {stats:?}");
    numeric(runtime.run_included_compilation(&mut vm,compiled,"include".into(),path.clone(),false,None).unwrap(),41);

    // Exercise a real built-in provider installation, then a new admitted
    // include carrying references to that exact installed identity. The parser
    // and dependency tracker select the production registration, not a test flag.
    let source="<?php $file=new SplFileInfo('/tmp/owned.php');return 1;";
    let path=write(&base,"install-library.php",source);let bundle=bundle_from_source(source,language,path.clone());
    let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&path).unwrap();
    numeric(runtime.run_included_compilation(&mut vm,compiled,"include".into(),path.clone(),false,None).unwrap(),1);
    let source="<?php $file=new SplFileInfo('/tmp/owned.php');class InstalledLibraryOwned{public function value(){return 7;}}return (new InstalledLibraryOwned)->value();";
    let path=write(&base,"reuse-library.php",source);let bundle=bundle_from_source(source,language,path.clone());
    crate::primitives::deferred_class_bodies::take_statistics();
    let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&path).unwrap();
    let stats=crate::primitives::deferred_class_bodies::take_statistics();
    assert!(stats.methods>0,"production owned class admission: {stats:?}");
    assert!(stats.installed_library_methods>0,"production installed library body elision: {stats:?}");
    numeric(runtime.run_included_compilation(&mut vm,compiled,"include".into(),path.clone(),false,None).unwrap(),7);

    // Installed library-reference headers alone must not select donation for
    // an otherwise functions-only unit.
    let source="<?php $file=new SplFileInfo('/tmp/owned.php');function owned_donation_plain(){return 8;}return owned_donation_plain();";
    let path=write(&base,"functions-only.php",source);let bundle=bundle_from_source(source,language,path.clone());
    crate::primitives::deferred_class_bodies::take_statistics();
    let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&path).unwrap();
    let stats=crate::primitives::deferred_class_bodies::take_statistics();
    assert_eq!(stats.methods,0);
    assert!(stats.installed_library_methods>0,"functions-only production metadata admission: {stats:?}");
    numeric(runtime.run_included_compilation(&mut vm,compiled,"include".into(),path.clone(),false,None).unwrap(),8);

    // Source declarations remain available to a later include and cache replay;
    // the trait unit and augmentation-using unit must decline donation.
    let source="<?php trait OwnedDonationServiceT{public function value(){return 42;}}return 1;";
    let path=write(&base,"trait.php",source);let bundle=bundle_from_source(source,language,path.clone());
    crate::primitives::deferred_class_bodies::take_statistics();
    let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&path).unwrap();
    assert_eq!(crate::primitives::deferred_class_bodies::take_statistics().methods,0);
    assert!(!compiled.source_declarations.is_empty());
    numeric(runtime.run_included_compilation(&mut vm,compiled,"include".into(),path.clone(),false,None).unwrap(),1);
    let source="<?php class OwnedDonationServiceB{use OwnedDonationServiceT;}return (new OwnedDonationServiceB)->value();";
    let path=write(&base,"use-trait.php",source);let bundle=bundle_from_source(source,language,path.clone());
    let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&path).unwrap();
    assert_eq!(crate::primitives::deferred_class_bodies::take_statistics().methods,0);
    numeric(runtime.run_included_compilation(&mut vm,compiled,"include".into(),path.clone(),false,None).unwrap(),42);

    // Fused completion deliberately retains its prepared AST for rejection and
    // cancellation fallback. This path MUST remain borrowed, even with the hook.
    let source="<?php class OwnedDonationFused{public function value(){return 9;}}return 9;";
    let path=write(&base,"fused.php",source);let pool=crate::preparation::PreparationPool::new(1,1).unwrap();
    let mut input=prepared(&pool,&path,language);let before=format!("{:?}",input.module);
    let context=FusedLoweringContext::capture(&vm,language,allow_completion,None,registration).unwrap();
    crate::primitives::deferred_class_bodies::take_statistics();
    let lowered=context.lower_prepared(&mut input,language,&Arc::new(AtomicUsize::new(0))).unwrap();
    assert!(lowered.compilation.is_ok());
    assert_eq!(crate::primitives::deferred_class_bodies::take_statistics().methods,0);
    assert_eq!(before,format!("{:?}",input.module),"fused fallback source AST changed");
    drop(lowered);

    // Borrowed fused completion may omit only installed-provider bodies;
    // it retains the complete original AST for foreground rejection fallback.
    let source="<?php function borrowed_fused_unused_provider(){return new SplFileInfo('/tmp/unused.php');}return 61;";
    let path=write(&base,"borrowed-installed-provider.php",source);
    let mut input=prepared(&pool,&path,language);let before=format!("{:?}",input.module);
    crate::primitives::deferred_class_bodies::take_statistics();
    let lowered=context.lower_prepared(&mut input,language,&Arc::new(AtomicUsize::new(0))).expect("installed-provider borrowed completion");
    let stats=crate::primitives::deferred_class_bodies::take_statistics();
    assert_eq!(stats.methods,0);assert!(stats.installed_library_methods>0,"{stats:?}");
    assert_eq!(before,format!("{:?}",input.module),"borrowed provider fallback AST changed");
    let bundle=bundle_from_source(source,language,path.clone());
    assert!(lowered.valid(&bundle,&vm),"real installed-provider certificate");
    numeric(runtime.run_included_compilation(&mut vm,lowered.compilation.unwrap().into_raw(),"include".into(),path.clone(),false,None).unwrap(),61);

    // Real separately queued worker owns its PreparedFile and consumes it once.
    let source="<?php class OwnedDonationWorker{public function value(){return 43;}}return (new OwnedDonationWorker)->value();";
    let path=write(&base,"worker.php",source);let ticket=pool.submit(path.clone(),language).unwrap();
    let deadline=std::time::Instant::now()+std::time::Duration::from_secs(30);
    while !ticket.is_ready(){assert!(std::time::Instant::now()<deadline,"worker preparation timeout");std::thread::yield_now();}
    runtime.preparing.insert(path.clone(),ticket);
    runtime.stage_ready_lowering_on(&pool,&vm,language,&path);
    let ticket=runtime.lowering.remove(&path).expect("actual worker lowering admission");
    let lowered=ticket.result.recv_timeout(std::time::Duration::from_secs(30)).expect("worker lowering result");
    let bundle=bundle_from_source(source,language,path.clone());
    assert!(lowered.valid(&bundle,&vm),"worker source/frontend/dependency certificate");
    numeric(runtime.run_included_compilation(&mut vm,lowered.compilation.unwrap().into_raw(),"include".into(),path.clone(),false,None).unwrap(),43);
}
