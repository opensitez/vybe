//! Compiler consumer for opt-in request code space. Private candidate.
use super::*;
use vybe_runtime::vm::LoadedRequestUnit;

#[derive(Clone)]
pub(super) struct SourceReceipt { observations:Vec<crate::bundle::SourceObservation> }
impl SourceReceipt {
    pub(super) fn capture(own:Option<crate::bundle::SourceObservation>,observations:Option<Vec<crate::bundle::SourceObservation>>)->Option<Self>{
        let mut observations=observations?;observations.push(own?);
        let result=Self{observations};result.valid().then_some(result)
    }
    pub(super) fn valid(&self)->bool{!self.observations.is_empty()&&self.observations.iter().all(|item|item.valid())}
}
#[derive(Clone)]
pub(super) struct Callable { pub(super) name:String, pub(super) arity:u8, pub(super) chunk:usize }
#[derive(Clone)]
pub(super) struct OwnedProvider {
    pub(super) name:String, pub(super) identity:String,
    object:std::sync::Weak<std::sync::Mutex<Object>>, chunk:usize,
}
impl OwnedProvider {
    fn capture(vm:&VM,unit:&LoadedRequestUnit,name:&str,identity:&str)->Option<Self> {
        let Value::Object(object)=vm.global(name).or_else(||vm.global(&name.to_lowercase()))? else{return None};
        let guard=object.lock().ok()?;
        let ObjectKind::Function(function)=&guard.kind else{return None};
        if !unit.chunk_range().contains(&function.chunk_index) {return None;}
        if !matches!(guard.properties.get(crate::primitives::reflection::FIELD_LIBRARY_IDENTITY),
            Some(Value::String(value)) if value.as_ref()==identity) {return None;}
        if !matches!(guard.properties.get(crate::primitives::reflection::FIELD_KIND),
            Some(Value::String(value)) if matches!(value.as_ref(),"class"|"interface")) {return None;}
        Some(Self{name:name.into(),identity:identity.into(),object:Arc::downgrade(object),chunk:function.chunk_index})
    }
    fn valid(&self,vm:&VM)->bool {
        let Some(expected)=self.object.upgrade() else{return false};
        let Some(Value::Object(actual))=vm.global(&self.name).or_else(||vm.global(&self.name.to_lowercase())) else{return false};
        if !Arc::ptr_eq(actual,&expected) {return false;}
        let Ok(object)=actual.lock() else{return false};
        matches!(&object.kind,ObjectKind::Function(f) if f.chunk_index==self.chunk)
            && matches!(object.properties.get(crate::primitives::reflection::FIELD_LIBRARY_IDENTITY),
                Some(Value::String(identity)) if identity.as_ref()==self.identity)
    }
}
#[derive(Clone)]
pub(super) struct ResidentEntry {
    pub(super) unit:LoadedRequestUnit, pub(super) source:SourceReceipt, frontend:FrontendIdentity,
    entry:PathBuf, pub(super) metadata:DynamicCompilation, callables:Vec<Callable>, pub(super) owned:Vec<OwnedProvider>,
}
impl ResidentEntry {
    pub(super) fn dependencies_valid(&self,vm:&VM)->bool {
        if !self.metadata.required_libraries.iter().all(|library|library_installed(vm,&library.name,&library.identity)) {return false;}
        let Some(deps)=&self.metadata.compile_dependencies else{#[cfg(feature="service-verification")]eprintln!("resident missing compile dependencies");return false};
        if (deps.profile_source)()!=deps.profile
            || !vybe_runtime::registry::find(deps.language_name).is_some_and(|l|FrontendIdentity::of(l)==deps.frontend)
            || !deps.certificate.lock().is_ok_and(|c|c.validates(&vm.modules)&&c.runtime_helpers_valid(vm)) {return false;}
        deps.libraries.iter().all(|(name,identity,expected)| {
            let installed=library_installed(vm,name,identity);
            installed==*expected || (!*expected && installed && self.owned.iter().any(|p|
                p.name==*name && p.identity==*identity && p.valid(vm)))
        })
    }
    pub(super) fn capture(vm:&VM,unit:LoadedRequestUnit,source:SourceReceipt,frontend:FrontendIdentity,
        entry:PathBuf,metadata:DynamicCompilation,callables:Vec<Callable>)->Option<Self> {
        // No full chunk graph may escape into request cache metadata.
        if !metadata.chunks.is_empty() || metadata.prepared_sources.is_some()
            || metadata.component_source || metadata.component_binary.is_some() || !source.valid() {
            #[cfg(feature="service-verification")]eprintln!("resident capture denied metadata: chunks={} prepared={} component={} source={}",metadata.chunks.len(),metadata.prepared_sources.is_some(),metadata.component_source,source.valid());return None;}
        let mut owned=Vec::new();
        if let Some(deps)=&metadata.compile_dependencies {
            for (name,identity,expected) in &deps.libraries {
                if !*expected && library_installed(vm,name,identity) {
                    let provider=OwnedProvider::capture(vm,&unit,name,identity);
                    #[cfg(feature="service-verification")]if provider.is_none(){eprintln!("resident owned-provider denied: {name}");}
                    owned.push(provider?);
                }
            }
        }
        let result=Self{unit,source,frontend,entry,metadata,callables,owned};
        let live=vm.request_unit_valid(&result.unit);let deps=result.dependencies_valid(vm);
        #[cfg(feature="service-verification")]if !live||!deps{eprintln!("resident capture denied authority: live={live} dependencies={deps}");}
        (live&&deps).then_some(result)
    }
    fn valid(&self,vm:&VM,language:Language,entry:&Path,origin:&Path)->bool {
        let frontend=self.frontend==FrontendIdentity::of(language);let entry=self.entry==entry;
        let origin=self.metadata.entry_path.as_deref()==Some(origin);let source=self.source.valid();let unit=vm.request_unit_valid(&self.unit);let deps=self.dependencies_valid(vm);
        #[cfg(feature="service-verification")]if !(frontend&&entry&&origin&&source&&unit&&deps){eprintln!("resident hit denied: frontend={frontend} entry={entry} origin={origin} source={source} unit={unit} deps={deps} audit={:?}",self.metadata.dependency_failure_for_audit(vm));}
        frontend&&entry&&origin&&source&&unit&&deps
    }
}
impl PhpIncludeRuntime {
    pub(super) fn resident_requested(&self,language:Language)->bool {
        crate::frontend_preparation::snapshot(language.parse).capabilities.request_code_space
    }
    pub(super) fn resident_hit(&mut self,vm:&VM,path:&Path,language:Language,entry:&Path,origin:&Path)->Option<ResidentHit> {
        if !self.resident_requested(language) {
            if !self.resident_units.is_empty(){self.resident_units.clear();}
            return None;
        }
        if self.resident_units.get(path).is_some_and(|e|!e.valid(vm,language,entry,origin)) {
            self.resident_units.remove(path);return None;
        }
        let hit=self.resident_units.get(path).map(clone_for_hit);
        hit
    }
    pub(super) fn run_resident_include(&mut self,vm:&mut VM,resident:ResidentHit,path:PathBuf,
        publishes_once:bool,scope:Option<Value>,imports:Vec<ImportTarget>)->Result<Value,String> {
        let caller_locals=if scope.is_some(){Vec::new()}else{vm.current_source_scope()
            .filter(|(name,_)|name!="<script>").map(|(_,locals)|locals).unwrap_or_default()};
        crate::host_imports::install(vm,&resident.metadata.host_imports);
        for declaration in &resident.callables {
            install_callable_global(vm,&declaration.name,declaration.arity,declaration.chunk,false);
        }
        self.execute_include_scope(vm,&resident.metadata.source_declarations,path,publishes_once,scope,
            caller_locals,|vm|vm.run_resident_nested(&resident.unit,imports))
    }
}

// Test-only owned variant reproduces the old clone without allocating an Arc
// around it. Normal builds use only the Arc alias and have no policy branch.
#[cfg(not(feature="service-verification"))]
pub(super) type ResidentHit=Arc<ResidentEntry>;
#[cfg(feature="service-verification")]
pub(super) enum ResidentHit { Shared(Arc<ResidentEntry>), Deep(ResidentEntry) }
#[cfg(feature="service-verification")]
impl std::ops::Deref for ResidentHit {
    type Target=ResidentEntry;
    fn deref(&self)->&ResidentEntry{match self{Self::Shared(entry)=>entry,Self::Deep(entry)=>entry}}
}
#[cfg(feature="service-verification")]
#[derive(Clone,Copy,Debug,Default,PartialEq,Eq)]
pub struct ResidentEntryCloneCounts {pub deep:usize,pub shared:usize}
#[cfg(feature="service-verification")]
thread_local! {
    static DEEP_RESIDENT_CONTROL:std::cell::Cell<bool>=const{std::cell::Cell::new(false)};
    static RESIDENT_CLONE_COUNTS:std::cell::Cell<ResidentEntryCloneCounts>=const{std::cell::Cell::new(ResidentEntryCloneCounts{deep:0,shared:0})};
}
#[cfg(feature="service-verification")]
pub fn with_resident_entry_clone_control<R>(deep:bool,run:impl FnOnce()->R)->(R,ResidentEntryCloneCounts){
    struct Restore(bool,ResidentEntryCloneCounts);
    impl Drop for Restore{fn drop(&mut self){DEEP_RESIDENT_CONTROL.set(self.0);RESIDENT_CLONE_COUNTS.set(self.1);}}
    let restore=Restore(DEEP_RESIDENT_CONTROL.replace(deep),RESIDENT_CLONE_COUNTS.replace(ResidentEntryCloneCounts::default()));
    let value=run();let counts=RESIDENT_CLONE_COUNTS.get();drop(restore);(value,counts)
}
fn clone_for_hit(entry:&Arc<ResidentEntry>)->ResidentHit {
    #[cfg(feature="service-verification")]{
        let deep=DEEP_RESIDENT_CONTROL.get();let mut counts=RESIDENT_CLONE_COUNTS.get();
        if deep{counts.deep+=1;}else{counts.shared+=1;}RESIDENT_CLONE_COUNTS.set(counts);
        if deep{ResidentHit::Deep((**entry).clone())}else{ResidentHit::Shared(Arc::clone(entry))}
    }
    #[cfg(not(feature="service-verification"))]{Arc::clone(entry)}
}
