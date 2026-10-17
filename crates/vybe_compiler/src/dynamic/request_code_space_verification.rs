//! Actual PHP service verifier. Included only by service-verification.
use super::*;
struct Restore(crate::frontend_preparation::Capabilities,vybe_runtime::registry::ParseFn);
impl Drop for Restore {fn drop(&mut self){crate::frontend_preparation::register(self.1,self.0);}}
struct Scratch(PathBuf);
impl Drop for Scratch {fn drop(&mut self){let _=fs::remove_dir_all(&self.0);}}
fn number(value:Value)->i32 {match value {Value::I32(n)=>n,Value::I64(n)=>n as i32,Value::F64(n)=>n as i32,other=>panic!("number required: {other:?}")}}

pub(super) fn verify(mut vm:VM) {
    let language=languages::find_by_name("php").unwrap();
    let registration=crate::frontend_preparation::snapshot(language.parse);
    let _restore=Restore(registration.capabilities,language.parse);
    let identity=FrontendIdentity::of(language);
    crate::frontend_preparation::register(language.parse,crate::frontend_preparation::Capabilities {
        request_code_space:!registration.capabilities.request_code_space,..registration.capabilities
    });
    assert_ne!(identity,FrontendIdentity::of(language),"request code-space bit invalidates producer identity");
    crate::frontend_preparation::register(language.parse,registration.capabilities);
    assert!(registration.capabilities.request_code_space,"actual generated PHP grammar must select request code space");
    let original=registration.identity();let mut ids=Vec::new();
    for setup in [false,true]{for request in [false,true]{
        crate::frontend_preparation::register(language.parse,crate::frontend_preparation::Capabilities{request_setup_reuse:setup,request_code_space:request,..registration.capabilities});
        let id=crate::frontend_preparation::snapshot(language.parse).identity();assert_eq!(id[9],original[9],"producer binding identity must remain unchanged");assert_eq!(id[2],setup as usize|((request as usize)<<1));assert!(!ids.contains(&id));ids.push(id);
    }}
    crate::frontend_preparation::register(language.parse,registration.capabilities);
    let base=Scratch(std::env::temp_dir().join(format!("vybe-request-code-{}",uuid::Uuid::new_v4())));
    fs::create_dir(&base.0).unwrap();
    let child=base.0.join("child.php");fs::write(&child,"<?php $x += 1; return $x;").unwrap();
    let caller=base.0.join("caller.php");
    let include=php_single_quoted_literal(&child.to_string_lossy());
    let source=format!("<?php function resident_first(){{$x=10;include {include};return $x;}}function resident_second(){{$x=20;include {include};return $x;}}return resident_first()*100+resident_second();");
    fs::write(&caller,source).unwrap();
    let mut service=RuntimeCompilerService::new(&mut vm);
    assert_eq!(number(service.compile_and_run_path(&caller).unwrap()),1121);
    assert_eq!(service.php_runtime.resident_hits,1,"second plain include must execute existing code");
    assert_eq!(service.php_runtime.resident_units.len(),1);
    let resident=service.php_runtime.resident_units.values().next().unwrap().clone();
    assert!(resident.metadata.chunks.is_empty(),"request cache must not retain a second Chunk graph");
    assert!(resident.metadata.source_declarations.is_empty());
    assert!(!resident.unit.chunk_range().is_empty());

    drop(service);

    let once_child=base.0.join("once-child.php");fs::write(&once_child,"<?php $n += 1; return $n;").unwrap();
    let path=php_single_quoted_literal(&once_child.to_string_lossy());
    fs::write(&caller,format!("<?php $n=0;include {path};require_once {path};include {path};return $n;")).unwrap();
    let mut service=RuntimeCompilerService::new(&mut vm);
    assert_eq!(number(service.compile_and_run_path(&caller).unwrap()),2);
    assert_eq!(service.php_runtime.resident_hits,1);

    // An edit at the same canonical source path invalidates the source receipt.
    let entry=service.php_runtime.resident_units.values().next().unwrap().clone();
    fs::write(&once_child,"<?php $n += 1000; return $n;").unwrap();
    assert!(!entry.source.valid());
    let active=service.php_runtime.activate(service.vm,None);drop(active);
    assert!(service.php_runtime.resident_units.is_empty(),"new request discards all live handles");
    assert!(service.vm.request_unit_valid(&entry.unit),"request cache clear need not invalidate unrelated VM code handles");
    drop(service);

    // A first include can itself install a library provider. Its exact
    // constructor body must come from this loaded unit; identity text alone
    // is insufficient. This fixture MUST prove that transition, not skip it.
    let library_child=base.0.join("library-child.php");
    fs::write(&library_child,"<?php return (new SplFileInfo('one.php'))->getFilename();").unwrap();
    let path=php_single_quoted_literal(&library_child.to_string_lossy());
    fs::write(&caller,format!("<?php $a=include {path};$b=include {path};return $a.$b;")).unwrap();
    let mut service=RuntimeCompilerService::new(&mut vm);
    let value=service.compile_and_run_path(&caller).unwrap();
    assert!(matches!(value,Value::String(ref value) if value.as_ref()=="one.phpone.php"));
    assert_eq!(service.php_runtime.resident_hits,1);
    let resident=service.php_runtime.resident_units.values().next().unwrap().clone();
    assert!(!resident.owned.is_empty(),"real newly installed library must have an exact owned-provider receipt");
    let provider=resident.owned.first().unwrap();
    let original=service.vm.global(&provider.name).or_else(||service.vm.global(&provider.name.to_lowercase())).unwrap().clone();
    let Value::Object(original)=original else{unreachable!()};
    let replacement=original.lock().unwrap().clone();
    let binding=if service.vm.global(&provider.name).is_some(){provider.name.clone()}else{provider.name.to_lowercase()};
    service.vm.set_global_owned(binding.clone(),Value::Object(vybe_runtime::heap::alloc(replacement)));
    assert!(library_installed(service.vm,&provider.name,&provider.identity),"identity text is intentionally unchanged");
    assert!(!resident.dependencies_valid(service.vm),"rebound provider must invalidate resident reuse");
    service.vm.set_global_owned(binding,Value::Object(original));
    drop(service);
    verify_capture_denial_retains_raw_cache(&mut vm,&base,language);
    verify_request_raw_retirement(&mut vm,&base,language);
    verify_source_observations(&base);
    verify_all_modes_and_classes(&mut vm,&base);
    verify_actual_edit(&mut vm,&base);
    println!("PASS actual PHP request code space: caller scope, once, edits, provider binding, raw fallback, source observations, namespaces/classes");
}

#[derive(Default)]
struct RawFallbackCache {
    request_only:bool,
    stamps:Mutex<HashMap<(PathBuf,u64),Vec<crate::bundle::SourceObservation>>>,
    entries:Mutex<HashMap<(PathBuf,u64),DynamicCompilation>>,
    hits:AtomicUsize, stores:AtomicUsize, retired:AtomicUsize,
}
impl IncludeCompileCache for RawFallbackCache {
    fn request_local_only(&self)->bool{self.request_only}
    fn retire_request_snapshot(&self,path:&Path,fingerprint:u64)->bool{
        assert!(self.request_only);
        let removed=self.entries.lock().unwrap().remove(&(path.to_path_buf(),fingerprint)).is_some();
        if removed{self.retired.fetch_add(1,Ordering::Relaxed);}removed
    }

    fn get(&self,path:&Path,fingerprint:u64)->Option<Result<DynamicCompilation,String>> {
        let key=(path.to_path_buf(),fingerprint);
        if self.stamps.lock().unwrap().get(&key).is_some_and(|s|!s.iter().all(|s|s.valid())){self.entries.lock().unwrap().remove(&key);return None;}
        let hit=self.entries.lock().unwrap().get(&key).cloned();
        if hit.is_some(){self.hits.fetch_add(1,Ordering::Relaxed);}hit.map(Ok)
    }
    fn store(&self,path:&Path,fingerprint:u64,deps:Option<Vec<PathBuf>>,result:Result<&DynamicCompilation,&str>) {
        if deps.is_none(){return;}
        if let Ok(compiled)=result {self.entries.lock().unwrap().insert((path.to_path_buf(),fingerprint),compiled.clone());self.stores.fetch_add(1,Ordering::Relaxed);}
    }
    fn store_with_chunks(&self,path:&Path,fingerprint:u64,deps:Option<Vec<PathBuf>>,
        result:Result<(&DynamicCompilation,&[Chunk]),&str>) {
        let Some(deps)=deps else{return};
        self.stamps.lock().unwrap().insert((path.to_path_buf(),fingerprint),deps.iter().map(|p|crate::bundle::SourceObservation::capture(p)).collect());
        if let Ok((metadata,chunks))=result {
            let mut compiled=metadata.clone();compiled.chunks=chunks.to_vec();
            self.entries.lock().unwrap().insert((path.to_path_buf(),fingerprint),compiled);
            self.stores.fetch_add(1,Ordering::Relaxed);
        }
    }

}
fn verify_capture_denial_retains_raw_cache(vm:&mut VM,base:&Scratch,language:Language) {
    // Warm ordinary library definitions so this test isolates epoch rejection,
    // not the independently tested negative provider transition.
    {
        let mut boot=RuntimeCompilerService::new(vm);
        let compiled=boot.compile_source("<?php return (new Exception('boot'))->getMessage();",language,base.0.join("boot.php")).unwrap();
        boot.run_compiled(compiled).unwrap();
    }
    let calls=Arc::new(AtomicUsize::new(0));let count=calls.clone();
    vm.register_host_fn("resident-test","invalidate",Box::new(move|_,_| {
        count.fetch_add(1,Ordering::Relaxed);
        ACTIVE_PHP_RUNTIME.with(|slot| {
            let runtime=slot.borrow().expect("active PHP runtime");
            let runtime=unsafe{&mut *runtime};
            assert!(!runtime.vm.is_null());
            unsafe{&mut *runtime.vm}.invalidate_request_units();
        });
        Value::Null
    }));
    let index=vm.host_registry[&("resident-test".to_string(),"invalidate".to_string())];
    let mut function=Object::new();function.kind=ObjectKind::HostFunction(index);
    let function=Value::Object(vybe_runtime::heap::alloc(function));
    vm.set_global("resident_break",function.clone());vm.set_global("__php_fn_resident_break",function);
    let child=base.0.join("capture-denied.php");fs::write(&child,"<?php resident_break();return 7;").unwrap();
    let path=php_single_quoted_literal(&child.to_string_lossy());
    let caller=base.0.join("capture-denial-caller.php");
    fs::write(&caller,format!("<?php $a=include {path};$b=include {path};return $a+$b;")).unwrap();
    let cache=Arc::new(RawFallbackCache{request_only:true,..Default::default()});let mut service=RuntimeCompilerService::new(vm);
    service.set_include_cache(cache.clone());
    assert_eq!(number(service.compile_and_run_path(&caller).unwrap()),14);
    assert_eq!(calls.load(Ordering::Relaxed),2,"script executes on both raw-cache loads");
    assert!(service.php_runtime.resident_units.is_empty(),"epoch invalidation must deny post-load resident capture");
    assert_eq!(cache.stores.load(Ordering::Relaxed),1,"first compilation keeps its ordinary raw snapshot");
    assert_eq!(cache.hits.load(Ordering::Relaxed),1,"second include reuses raw compilation after capture denial");
    assert_eq!(cache.retired.load(Ordering::Relaxed),0);
    assert_eq!(cache.entries.lock().unwrap().len(),1);
}

fn verify_source_observations(base:&Scratch){
    let path=base.0.join("negative-dependency.php");
    let own=crate::bundle::SourceObservation::capture(&base.0.join("caller.php"));
    let (_,observations)=crate::bundle::record_source_observations(||{crate::bundle::note_source_read(&path);assert!(fs::read_to_string(&path).is_err());});
    let receipt=SourceReceipt::capture(Some(own),observations).expect("attributable missing dependency");assert!(receipt.valid());
    fs::write(&path,"<?php return 77;").unwrap();assert!(!receipt.valid(),"new negative dependency invalidates");
    let own=crate::bundle::SourceObservation::capture(&path);
    let (_,observations)=crate::bundle::record_source_observations(||{crate::bundle::note_source_read(&path);let _=fs::read_to_string(&path).unwrap();fs::write(&path,"<?php return 999999;").unwrap();});
    assert!(SourceReceipt::capture(Some(own),observations).is_none(),"change during compile denies capture");
    let (_,outer)=crate::bundle::record_source_observations(||{
        let (_,inner)=crate::bundle::record_source_observations(||crate::bundle::note_source_read(&path));assert!(inner.is_none());
    });assert!(outer.is_some());
}
fn verify_all_modes_and_classes(vm:&mut VM,base:&Scratch){
    let child=base.0.join("all-modes.php");fs::write(&child,"<?php $n += 1; return $n;").unwrap();let p=php_single_quoted_literal(&child.to_string_lossy());
    let caller=base.0.join("modes-caller.php");fs::write(&caller,format!("<?php $n=0;include {p};require {p};include_once {p};require_once {p};include {p};return $n;")).unwrap();
    let mut service=RuntimeCompilerService::new(vm);assert_eq!(number(service.compile_and_run_path(&caller).unwrap()),3);assert_eq!(service.php_runtime.resident_hits,2);
    let entry=service.php_runtime.resident_units.values().next().unwrap();assert_eq!(entry.unit.chunk_range().len(),service.vm.chunks.len()-entry.unit.chunk_range().start,"no repeated code append");drop(service);
    let class=base.0.join("source-class.php");fs::write(&class,"<?php namespace ResidentSpace; class Box { public function value(){return 19;} } return 19;").unwrap();
    let use_class=base.0.join("use-class.php");fs::write(&use_class,"<?php return (new \\ResidentSpace\\Box())->value();").unwrap();
    let c=php_single_quoted_literal(&class.to_string_lossy());let u=php_single_quoted_literal(&use_class.to_string_lossy());
    fs::write(&caller,format!("<?php require_once {c};require_once {c};$a=include {u};$b=include {u};return $a+$b;")).unwrap();
    let mut service=RuntimeCompilerService::new(vm);assert_eq!(number(service.compile_and_run_path(&caller).unwrap()),38);assert_eq!(service.php_runtime.resident_hits,1);
    let canonical=class.canonicalize().unwrap();assert!(service.php_runtime.resident_units.contains_key(&canonical),"actual namespaced source class must be admitted");
}

fn verify_actual_edit(vm:&mut VM,base:&Scratch){
    let child=base.0.join("edited.php");fs::write(&child,"<?php return 7;").unwrap();let edit=child.clone();
    vm.register_host_fn("resident-test","edit-source",Box::new(move|_,_|{fs::write(&edit,"<?php return 9999;").unwrap();Value::Null}));
    let index=vm.host_registry[&("resident-test".into(),"edit-source".into())];let mut function=Object::new();function.kind=ObjectKind::HostFunction(index);let function=Value::Object(vybe_runtime::heap::alloc(function));vm.set_global("resident_edit",function.clone());vm.set_global("__php_fn_resident_edit",function);
    let caller=base.0.join("edit-caller.php");let p=php_single_quoted_literal(&child.to_string_lossy());fs::write(&caller,format!("<?php $a=include {p};resident_edit();$b=include {p};return $a+$b;")).unwrap();
    let mut service=RuntimeCompilerService::new(vm);assert_eq!(number(service.compile_and_run_path(&caller).unwrap()),10006);assert_eq!(service.php_runtime.resident_hits,0,"edited source must compile/load new code");
}

fn verify_request_raw_retirement(vm:&mut VM,base:&Scratch,language:Language){
    let child=base.0.join("retire-child.php");let caller=base.0.join("retire-caller.php");
    let path=php_single_quoted_literal(&child.to_string_lossy());
    let original=crate::frontend_preparation::snapshot(language.parse).capabilities;
    // Legacy backends and unannotated/default frontends must never retire.
    for(request_only,optin)in[(false,true),(true,false),(true,true)]{
        crate::frontend_preparation::register(language.parse,crate::frontend_preparation::Capabilities{request_code_space:optin,..original});
        fs::write(&child,"<?php $n+=1;return $n;").unwrap();
        fs::write(&caller,format!("<?php $n=0;include {path};include {path};include {path};return $n;")).unwrap();
        let cache=Arc::new(RawFallbackCache{request_only,..Default::default()});
        let mut service=RuntimeCompilerService::new(vm);service.set_include_cache(cache.clone());
        assert_eq!(number(service.compile_and_run_path(&caller).unwrap()),3);
        assert_eq!(cache.stores.load(Ordering::Relaxed),1);
        let raw_bytes:usize=cache.entries.lock().unwrap().values().flat_map(|c|&c.chunks).map(|c|c.code.len()).sum();
        if request_only&&optin{
            assert_eq!(cache.retired.load(Ordering::Relaxed),1);assert!(cache.entries.lock().unwrap().is_empty());assert_eq!(raw_bytes,0);assert_eq!(service.php_runtime.resident_hits,2);
        }else{
            assert_eq!(cache.retired.load(Ordering::Relaxed),0);assert_eq!(cache.entries.lock().unwrap().len(),1);assert!(raw_bytes>0);
        }
        println!("retirement request_only={request_only} optin={optin} raw_code_bytes={raw_bytes} retired={} effects=3",cache.retired.load(Ordering::Relaxed));
    }
    crate::frontend_preparation::register(language.parse,original);
    // Edit during execution invalidates first receipt; second include recompiles
    // once, then only its newly valid snapshot can retire. No repeated effects.
    let edit=child.clone();vm.register_host_fn("retirement-test","edit",Box::new(move|_,_|{fs::write(&edit,"<?php $n+=100;return $n;").unwrap();Value::Null}));
    let index=vm.host_registry[&("retirement-test".into(),"edit".into())];let mut f=Object::new();f.kind=ObjectKind::HostFunction(index);let f=Value::Object(vybe_runtime::heap::alloc(f));vm.set_global("retirement_edit",f.clone());vm.set_global("__php_fn_retirement_edit",f);
    fs::write(&child,"<?php $n+=1;retirement_edit();return $n;").unwrap();
    fs::write(&caller,format!("<?php $n=0;include {path};include {path};include {path};return $n;")).unwrap();
    let cache=Arc::new(RawFallbackCache{request_only:true,..Default::default()});let mut service=RuntimeCompilerService::new(vm);service.set_include_cache(cache.clone());
    assert_eq!(number(service.compile_and_run_path(&caller).unwrap()),201);
    assert_eq!(cache.stores.load(Ordering::Relaxed),2);assert_eq!(cache.retired.load(Ordering::Relaxed),1);assert_eq!(service.php_runtime.resident_hits,1);assert!(cache.entries.lock().unwrap().is_empty());
    println!("retirement source-edit recompiles=2 retired=1 effects=201");
}
