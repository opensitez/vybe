//! Real request-local service verification, with the registered PHP frontend.
use super::*;
fn register_php() {
    #[cfg(test)]
    vybe_language_php::register();
    assert!(crate::languages::find_by_name("php").is_some(), "register PHP before running the service verification");
}
fn configured_vm() -> VM {
    let mut vm=VM::new();
    crate::primitives::platforms::register_platforms_all(&mut vm);
    vm
}
fn assert_numeric_value(value:Value,expected:f64) {
    let actual=match value {Value::I32(n)=>n as f64,Value::I64(n)=>n as f64,Value::F64(n)=>n,other=>panic!("expected numeric value, got {other:?}")};
    assert_eq!(actual,expected);
}
    fn retired_compiled_fixture(runtime:&mut super::PhpIncludeRuntime,vm:&mut VM,path:&Path,source:&str)->(super::LoweredSource,super::Bundle) {
        let language=crate::languages::find_by_name("php").unwrap();
        let bundle=super::bundle_from_source(source,language,path.to_path_buf());
        let pending=runtime.compile_pending_with_reuse(vm,&bundle,path,true).unwrap();
        assert!(pending.compilation.source_declarations.is_empty());
        let pending=super::PendingCompilation::prepare(pending.into_raw(),&runtime.link_budget);
        (super::LoweredSource {elapsed:Default::default(),path:path.to_path_buf(),source:source.to_string(),compilation:Ok(pending)},bundle)
    }

    #[cfg_attr(test, test)]
    pub(super) fn retired_compiled_service_hit_moves_code_releases_metadata_and_consumes_once() {
        register_php();
        let mut vm=configured_vm();let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        let path=std::env::temp_dir().join("retired-compiled-hit.php");
        let (lowered,bundle)=retired_compiled_fixture(&mut runtime,&mut vm,&path,"<?php return 7;");
        let expected=lowered.compilation.as_ref().unwrap().code.chunks()[0].code.clone();
        let ptr=lowered.compilation.as_ref().unwrap().code.chunks()[0].code.as_ptr();
        assert!(runtime.link_budget.load(std::sync::atomic::Ordering::Relaxed)>0);
        assert!(runtime.retired_compiled.retain(path.clone(),lowered));
        assert_eq!(runtime.link_budget.load(std::sync::atomic::Ordering::Relaxed),0);
        let pending=runtime.compile_pending_with_reuse(&mut vm,&bundle,&path,true).unwrap();
        assert_eq!(pending.code.chunks()[0].code.as_ptr(),ptr,"must consume the moved retained buffer");
        assert_eq!(pending.code.chunks()[0].code,expected);
        assert!(runtime.retired_compiled.entries.is_empty());
        assert_eq!(runtime.retired_compiled.payload_bytes,0);
        // A later plain include receives no cached mutable/relocated result.
        assert!(runtime.retired_compiled.take(&path).is_none());
        let compiled=pending.into_raw();
        let mut service=RuntimeCompilerService::new(&mut vm);
        assert_numeric_value(service.run_compiled(compiled).unwrap(),7.0);
    }

    #[cfg_attr(test, test)]
    pub(super) fn retired_compiled_service_revalidates_exact_source_and_negative_library_fact() {
        register_php();
        let language=crate::languages::find_by_name("php").unwrap();
        let mut vm=configured_vm();let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        let path=std::env::temp_dir().join("retired-compiled-source.php");
        let (lowered,_)=retired_compiled_fixture(&mut runtime,&mut vm,&path,"<?php return 7;");
        assert!(runtime.retired_compiled.retain(path.clone(),lowered));
        let changed=super::bundle_from_source("<?php return 8;",language,path.clone());
        let selected=runtime.compile_pending_with_reuse(&mut vm,&changed,&path,true).unwrap().into_raw();
        let control=runtime.compile_dynamic_php_with_reuse(&mut vm,&changed,&path,false).unwrap();
        assert_eq!(selected.chunks[0].code,control.chunks[0].code);
        assert!(runtime.retired_compiled.entries.is_empty());
        let (mut lowered,bundle)=retired_compiled_fixture(&mut runtime,&mut vm,&path,"<?php return 7;");
        let pending=lowered.compilation.as_mut().unwrap();
        let old=pending.compilation.compile_dependencies.as_ref().unwrap();
        pending.compilation.compile_dependencies=Some(Arc::new(super::CompilationDependencies {
            certificate:old.certificate.clone(),libraries:vec![("retiredguard".into(),"identity-v1".into(),false)],
            profile:old.profile,profile_source:old.profile_source,language_name:old.language_name,frontend:old.frontend,
        }));
        assert!(lowered.valid(&bundle,&vm));
        assert!(runtime.retired_compiled.retain(path.clone(),lowered));
        let mut object=vybe_runtime::value::Object::new();
        object.properties.insert(crate::primitives::reflection::FIELD_LIBRARY_IDENTITY.into(),Value::String(Arc::from("identity-v1")));
        vm.set_global("retiredguard",Value::Object(vybe_runtime::heap::alloc(object)));
        assert!(!runtime.retired_compiled.take(&path).unwrap().valid(&bundle,&vm));
    }

    #[cfg_attr(test, test)]
    pub(super) fn retired_compiled_count_payload_bounds_and_cancellation_release() {
        register_php();
        let mut vm=configured_vm();let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        for n in 0..=super::RETIRED_COMPILED_COUNT {
            let path=std::env::temp_dir().join(format!("retired-compiled-{n}.php"));
            let (lowered,_)=retired_compiled_fixture(&mut runtime,&mut vm,&path,"<?php return 7;");
            assert!(runtime.retired_compiled.retain(path,lowered));
        }
        assert_eq!(runtime.retired_compiled.entries.len(),super::RETIRED_COMPILED_COUNT);
        assert!(runtime.retired_compiled.entries.front().unwrap().key.ends_with("retired-compiled-1.php"));
        assert!(runtime.retired_compiled.payload_bytes<=super::RETIRED_COMPILED_PAYLOAD_BUDGET);
        let path=std::env::temp_dir().join("retired-compiled-oversize.php");
        let (mut lowered,_)=retired_compiled_fixture(&mut runtime,&mut vm,&path,"<?php return 7;");
        let pending=lowered.compilation.unwrap().into_raw();
        let mut raw=pending;raw.chunks[0].code.reserve(super::RETIRED_COMPILED_PAYLOAD_BUDGET+1);
        lowered.compilation=Ok(super::PendingCompilation::raw(raw));
        assert!(!runtime.retired_compiled.retain(path,lowered));
        assert_eq!(runtime.link_budget.load(std::sync::atomic::Ordering::Relaxed),0);
        // Two individually eligible ~9MiB buffers exceed the combined payload
        // limit: FIFO eviction must leave only the second, independent of count.
        runtime.retired_compiled.clear();
        runtime.retired_compiled.select(Some(super::FrontendIdentity::of(crate::languages::find_by_name("php").unwrap())));
        for n in 0..2 {
            let path=std::env::temp_dir().join(format!("retired-compiled-large-{n}.php"));
            let (mut lowered,_)=retired_compiled_fixture(&mut runtime,&mut vm,&path,"<?php return 7;");
            let mut raw=lowered.compilation.unwrap().into_raw();
            raw.chunks[0].code.reserve(9*1024*1024);
            lowered.compilation=Ok(super::PendingCompilation::raw(raw));
            assert!(runtime.retired_compiled.retain(path,lowered));
        }
        assert_eq!(runtime.retired_compiled.entries.len(),1);
        assert!(runtime.retired_compiled.entries.front().unwrap().key.ends_with("retired-compiled-large-1.php"));
        assert!(runtime.retired_compiled.payload_bytes<=super::RETIRED_COMPILED_PAYLOAD_BUDGET);
        runtime.cancel_request_preparation();
        assert!(runtime.retired_compiled.entries.is_empty());assert_eq!(runtime.retired_compiled.payload_bytes,0);
    }

    #[cfg_attr(test, test)]
    pub(super) fn retired_compiled_errors_are_not_retained_or_published() {
        register_php();
        let language=crate::languages::find_by_name("php").unwrap();
        let mut bank=super::RetiredCompiledBank::default();bank.select(Some(super::FrontendIdentity::of(language)));
        let path=PathBuf::from("unreached-retired-error.php");
        assert!(!bank.retain(path.clone(),super::LoweredSource {elapsed:Default::default(),path,source:"<?php function (".into(),compilation:Err("private speculative error".into())}));
        assert!(bank.entries.is_empty());assert_eq!(bank.payload_bytes,0);
    }

