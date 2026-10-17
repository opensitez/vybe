//! Exact-parser opt-in for direct statement generation into the shared emitter.
//! Bindings contain immutable hooks only. Opaque semantic facts belong to one
//! prepared source and travel with its original source/manifest, never globally.
use std::any::Any;
use std::path::Path;
use std::sync::{Arc,Mutex,OnceLock};
use vybe_ast::{Module,Statement};
use vybe_runtime::registry::ParseFn;

pub type StatementSink<'a> = dyn FnMut(usize,Statement)->Result<(),String>+'a;
pub type ProducerFn = fn(&str,Module,&mut StatementSink<'_>)->Result<(),String>;
/// Request-owned immutable semantic facts. Providers must not retain live VM
/// state or mutate these facts between preparation and producer replay.
pub type ProducerContext = Arc<dyn Any+Send+Sync>;
pub type PrepareWithContextFn = fn(&str,&Path,&mut crate::frontend_preparation::PreparationSink<'_>)
    ->Result<(crate::frontend_preparation::FrontendPreparation,ProducerContext),String>;
pub type ProduceWithContextFn = fn(&str,Module,&ProducerContext,&mut StatementSink<'_>)->Result<(),String>;
pub struct PreparedHooks {pub prepare:PrepareWithContextFn,pub produce:ProduceWithContextFn}
/// Source/grammar validation only: this hook never receives the compiler-owned
/// manifest and therefore cannot replace statements after semantic preparation.
pub type ValidateOwnedFn = fn(&str,&ProducerContext)->Result<(),String>;
pub struct PreparedOwnedHooks {pub prepare:PrepareWithContextFn,pub validate:ValidateOwnedFn}
#[derive(Clone,Copy)]
enum Hooks { Legacy(ProducerFn), Prepared(&'static PreparedHooks), PreparedOwned(&'static PreparedOwnedHooks) }
/// An immutable registry selection. Its id is never reused for a replacement;
/// a source prepared before re-registration keeps its original hooks/identity.
#[derive(Clone,Copy)]
pub(crate) struct Binding {id:usize,hooks:Hooks}
impl Binding {
    pub(crate) fn identity(self)->usize{self.id}
    pub(crate) fn prepare(self)->Option<PrepareWithContextFn>{match self.hooks{Hooks::Prepared(h)=>Some(h.prepare),Hooks::PreparedOwned(h)=>Some(h.prepare),_=>None}}
    pub(crate) fn capture(self,source:&str,context:Option<ProducerContext>)->Option<Production>{
        if matches!(self.hooks,Hooks::Prepared(_)|Hooks::PreparedOwned(_)) && context.is_none(){return None;}
        Some(Production{source:Arc::from(source),binding:self,context})
    }
}
struct Registry {next:usize,entries:Vec<(ParseFn,Binding)>}
fn entries()->&'static Mutex<Registry>{
    static ENTRIES:OnceLock<Mutex<Registry>>=OnceLock::new();
    ENTRIES.get_or_init(||Mutex::new(Registry{next:1,entries:Vec::new()}))
}
fn publish(parser:ParseFn,hooks:Hooks){
    let mut registry=entries().lock().unwrap();
    let index=registry.entries.iter().position(|entry|std::ptr::fn_addr_eq(entry.0,parser));
    if let Some(index)=index {
        let same=match(registry.entries[index].1.hooks,hooks){
            (Hooks::Legacy(a),Hooks::Legacy(b))=>std::ptr::fn_addr_eq(a,b),
            (Hooks::Prepared(a),Hooks::Prepared(b))=>std::ptr::eq(a,b),
            (Hooks::PreparedOwned(a),Hooks::PreparedOwned(b))=>std::ptr::eq(a,b),_=>false};
        if same{return;}
    }
    let id=registry.next;registry.next=id.checked_add(1).expect("statement binding identity exhausted");
    let binding=Binding{id,hooks};
    if let Some(index)=index{registry.entries[index].1=binding;}else{registry.entries.push((parser,binding));}
}
/// Default PEG parsers remain unregistered. Repeated bootstrap of identical
/// hooks preserves identity; an explicit replacement gets a new identity.
pub fn register(parser:ParseFn,producer:ProducerFn){publish(parser,Hooks::Legacy(producer));}
pub fn register_prepared(parser:ParseFn,hooks:&'static PreparedHooks){publish(parser,Hooks::Prepared(hooks));}
/// Opt in to compiler-owned statement consumption. Legacy prepared callbacks
/// retain their independent producer verification and do not select this path.
pub fn register_prepared_owned(parser:ParseFn,hooks:&'static PreparedOwnedHooks){publish(parser,Hooks::PreparedOwned(hooks));}
pub(crate) fn selected(parser:ParseFn)->Option<Binding>{
    entries().lock().unwrap().entries.iter().find(|entry|std::ptr::fn_addr_eq(entry.0,parser)).map(|entry|entry.1)
}
/// Source is the exact expanded/normalized text that produced this context.
/// Fields are private so a context cannot be detached/rebound to another source
/// by a later registry change. The manifest must remain that source's prepared
/// normalized manifest; source-path/provider binding may refine it in place.
#[derive(Clone)]
pub struct Production {source:Arc<str>,binding:Binding,context:Option<ProducerContext>}
impl Production {
    pub(crate) fn capture(parser:ParseFn,source:&str)->Option<Self>{selected(parser)?.capture(source,None)}
    pub(crate) fn is_owned(&self)->bool{matches!(self.binding.hooks,Hooks::PreparedOwned(_))}
    pub(crate) fn validate_owned(self)->Result<(),String>{
        match self.binding.hooks {
            Hooks::PreparedOwned(hooks)=>(hooks.validate)(&self.source,
                self.context.as_ref().ok_or("prepared owned context missing")?),
            _=>Err("statement production does not authorize owned consumption".into()),
        }
    }
    pub fn emit(self,module:Module,sink:&mut StatementSink<'_>)->Result<(),String>{
        match self.binding.hooks {
            Hooks::Legacy(produce)=>produce(&self.source,module,sink),
            Hooks::Prepared(hooks)=>(hooks.produce)(&self.source,module,
                self.context.as_ref().ok_or("prepared statement context missing")?,sink),
            Hooks::PreparedOwned(hooks)=>{
                (hooks.validate)(&self.source,self.context.as_ref().ok_or("prepared owned context missing")?)?;
                // Public callback consumers still use their ordinary checked
                // sink; only Bundle owns the private hash-free iterator path.
                for(index,statement)in module.body.into_iter().enumerate(){sink(index,statement)?;}
                Ok(())
            }
        }
    }
}

#[cfg(test)]
mod owned_binding_tests {
    use super::*;
    fn module(source:&str)->Module{Module{name:source.into(),language:vybe_ast::Lang::Unknown,
        body:vec![],imports:vec![],directives:Default::default(),canon:Default::default()}}
    fn prepare(source:&str,_:&Path,_:&mut crate::frontend_preparation::PreparationSink<'_>)
        ->Result<(crate::frontend_preparation::FrontendPreparation,ProducerContext),String>{
        Ok((crate::frontend_preparation::FrontendPreparation{module:module(source),hints:None},Arc::new(source.to_owned())))
    }
    fn validate(source:&str,context:&ProducerContext)->Result<(),String>{
        if context.downcast_ref::<String>().is_some_and(|original|original==source){Ok(())}
        else{Err("owned source/commit mismatch".into())}
    }
    static HOOKS:PreparedOwnedHooks=PreparedOwnedHooks{prepare,validate};
    #[test]
    fn owned_requires_context_and_repeated_registration_preserves_identity(){
        fn parser(source:&str)->Result<Module,String>{Ok(module(source))}
        register_prepared_owned(parser,&HOOKS);let binding=selected(parser).unwrap();
        register_prepared_owned(parser,&HOOKS);assert_eq!(selected(parser).unwrap().identity(),binding.identity());
        assert!(binding.prepare().is_some());assert!(binding.capture("A",None).is_none());
        let production=binding.capture("A",Some(Arc::new("A".to_owned()))).unwrap();
        assert!(production.is_owned());production.validate_owned().unwrap();
    }
    #[test]
    fn source_mismatch_rejects_before_public_sink_and_keeps_captured_binding(){
        fn parser(source:&str)->Result<Module,String>{Ok(module(source))}
        fn replacement(_: &str,_:Module,_:&mut StatementSink<'_>)->Result<(),String>{Err("replacement".into())}
        register_prepared_owned(parser,&HOOKS);let binding=selected(parser).unwrap();
        let captured=binding.capture("A",Some(Arc::new("A".to_owned()))).unwrap();
        register(parser,replacement);captured.validate_owned().unwrap();
        let wrong=binding.capture("A",Some(Arc::new("B".to_owned()))).unwrap();let mut called=false;
        assert!(wrong.emit(module("A"),&mut |_,_|{called=true;Ok(())}).unwrap_err().contains("mismatch"));assert!(!called);
    }
    #[test]
    fn legacy_registration_never_authorizes_owned_consumption(){
        fn parser(source:&str)->Result<Module,String>{Ok(module(source))}
        fn legacy(_: &str,_:Module,_:&mut StatementSink<'_>)->Result<(),String>{Ok(())}
        register(parser,legacy);let production=selected(parser).unwrap().capture("A",None).unwrap();
        assert!(!production.is_owned());assert!(production.validate_owned().is_err());
        fn unrelated(source:&str)->Result<Module,String>{Ok(module(source))}
        assert!(selected(unrelated).is_none());
    }
}
