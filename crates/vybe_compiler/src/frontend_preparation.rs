//! Explicit frontend construction receipts, produced by the same parse as the AST.
//! No grammar-engine dependency, execution permission, or filesystem resolution.
use std::path::Path;
use std::sync::{Mutex, OnceLock};
use vybe_ast::Module;
use vybe_runtime::registry::ParseFn;
use crate::preparation::DeferredPreparationHints;

pub struct FrontendPreparation {
    pub module: Module,
    /// Some (including empty) replaces post-parse AST hint discovery. These
    /// request-owned summaries schedule preparation only; actual loads validate
    /// current source, paths and caller scope through the existing include API.
    pub hints: Option<Box<dyn DeferredPreparationHints>>,
}
pub type ResolvePathFn=fn(&Path,&Path,&str,bool)->std::path::PathBuf;
pub type PreparationSink<'a>= dyn FnMut(Box<dyn DeferredPreparationHints>)+'a;
/// Providers publish one complete committed batch through the sink, or return
/// it in `hints`, never both. Protocol violations retain the AST but re-enable
/// legacy discovery; only the first batch is forwarded to bounded workers.
pub type PrepareFn = fn(&str, &Path, &mut PreparationSink<'_>) -> Result<FrontendPreparation, String>;
#[derive(Clone, Copy)]
pub struct FrontendPreparer { parser: ParseFn, prepare: PrepareFn, pub early_hints: bool, producer:Option<crate::statement_codegen::Binding> }
impl FrontendPreparer {
    pub fn matches(self, other:Self)->bool {
        std::ptr::fn_addr_eq(self.parser,other.parser)
            && std::ptr::fn_addr_eq(self.prepare,other.prepare) && self.early_hints==other.early_hints
            && self.producer.map(|p|p.identity())==other.producer.map(|p|p.identity())
    }
    pub(crate) fn matches_parser(self,parser:ParseFn)->bool{std::ptr::fn_addr_eq(self.parser,parser)}
    pub(crate) fn invoke(self,source:&str,caller:&Path,sink:&mut PreparationSink<'_>)->Result<FrontendPreparation,String>{(self.prepare)(source,caller,sink)}
}
/// Exact-parser opt-in capabilities owned by the compiler, not its runtime ABI.
#[derive(Clone,Copy,Default)]
pub struct Capabilities {
    pub prepare:Option<PrepareFn>,
    pub early_hints:bool,
    pub request_setup_reuse:bool,
    /// Explicit grammar code-space capability; absent grammars preserve unit linking.
    pub request_code_space:bool,
    pub helper_definitions:&'static [crate::primitives::helper_templates::Definition],
    pub argument_emitters:&'static [crate::primitives::argument_emission::Definition],
    pub resolve_path:Option<ResolvePathFn>,
    pub deferred_class_bodies:Option<crate::primitives::deferred_class_bodies::NormalizeFn>,
}
fn registrations()->&'static Mutex<Vec<(ParseFn,Capabilities)>>{
    static REGISTRY:OnceLock<Mutex<Vec<(ParseFn,Capabilities)>>>=OnceLock::new();
    REGISTRY.get_or_init(||Mutex::new(Vec::new()))
}
/// No request state belongs to a registration. An exact selected parser is
/// required; public PEG reference parsers remain unregistered by default.
pub fn register(parser:ParseFn,capabilities:Capabilities){
    let mut entries=registrations().lock().unwrap();
    if let Some(entry)=entries.iter_mut().find(|entry|std::ptr::fn_addr_eq(entry.0,parser)){entry.1=capabilities;}
    else{entries.push((parser,capabilities));}
}
/// Publish language bootstrap defaults without replacing an explicit selection.
/// Reinitializing a VM may initialize its language plugins again. Such setup
/// must not erase the immutable compiler policy already selected for a parser.
/// Use `register` when deliberately replacing that policy, including disabling
/// capabilities. The presence check and publication share the registry lock.
pub fn register_default(parser:ParseFn,capabilities:Capabilities){
    let mut entries=registrations().lock().unwrap();
    if !entries.iter().any(|entry|std::ptr::fn_addr_eq(entry.0,parser)){
        entries.push((parser,capabilities));
    }
}
pub fn capabilities(parser:ParseFn)->Capabilities {
    registrations().lock().unwrap().iter().find(|entry|std::ptr::fn_addr_eq(entry.0,parser))
        .map_or_else(Capabilities::default,|entry|entry.1)
}
/// A snapshot owns preparation policy and the immutable producer binding;
/// later parsing and cache identity never consult a replacement producer.
#[derive(Clone,Copy)]
pub struct Registration { parser:ParseFn, pub capabilities:Capabilities, producer:Option<crate::statement_codegen::Binding> }
impl Registration {
    pub fn preparer(self)->Option<FrontendPreparer>{self.capabilities.prepare.map(|prepare|
        FrontendPreparer{parser:self.parser,prepare,early_hints:self.capabilities.early_hints,producer:self.producer})}
    pub fn helper_selection(self)->crate::primitives::helper_templates::Selection{
        crate::primitives::helper_templates::Selection{parser:self.parser,definitions:self.capabilities.helper_definitions}
    }
    pub fn argument_selection(self)->crate::primitives::argument_emission::Selection {
        crate::primitives::argument_emission::Selection::new(self.parser, self.capabilities.argument_emitters)
    }
    pub(crate) fn identity(self)->[usize;10]{let entry=self.capabilities;
        [entry.prepare.map_or(0,|f|f as usize),entry.early_hints as usize,
         (entry.request_setup_reuse as usize)|((entry.request_code_space as usize)<<1),entry.resolve_path.map_or(0,|f|f as usize),
         if entry.helper_definitions.is_empty(){0}else{entry.helper_definitions.as_ptr() as usize},entry.helper_definitions.len(),
         if entry.argument_emitters.is_empty(){0}else{entry.argument_emitters.as_ptr() as usize},entry.argument_emitters.len(),entry.deferred_class_bodies.map_or(0,|f|f as usize),self.producer.map_or(0,|binding|binding.identity())]}
}
pub fn snapshot(parser:ParseFn)->Registration{Registration{parser,capabilities:capabilities(parser),producer:crate::statement_codegen::selected(parser)}}
pub fn for_parser(parser:ParseFn)->Option<FrontendPreparer>{snapshot(parser).preparer()}
pub(crate) fn identity(parser:ParseFn)->[usize;10]{snapshot(parser).identity()}
pub fn php_source_path(entry:&Path,caller:&Path,target:&str,fused:bool)->std::path::PathBuf{
    crate::dynamic::resolve_prepared_php_include_path(entry,caller,target,fused)
}
/// Resolve only parser-committed operand text, never the original source using
/// normalized offsets. Unsupported/computed expressions are not evaluated.
/// This PHP policy adapter is scheduling-only; the generic carrier is neutral.
pub fn php_operand_hints(operands:impl IntoIterator<Item=String>,caller:&Path)->Box<dyn DeferredPreparationHints>{
    Box::new(crate::php_include_hints::IncludeHints::from_operands(operands,caller))
}
pub fn matches_current(parser:ParseFn,prepared:Option<FrontendPreparer>)->bool{
    match (for_parser(parser),prepared){(None,None)=>true,(Some(a),Some(b))=>a.matches(b),_=>false}
}
/// Retain ownership through the bundle's ordinary import resolution and entry
/// rewriting. Legacy Module-returning APIs may discard receipts deliberately.
pub struct PreparedModule {
    pub module:Module,
    pub hints:Option<Box<dyn DeferredPreparationHints>>,
    pub preparer:Option<FrontendPreparer>,
    pub preparation_hints_complete:bool,
    pub statement_production:Option<crate::statement_codegen::Production>,
}
pub(crate) fn parse(source:&str,caller:&Path,parser:ParseFn)->Result<PreparedModule,String>{
    parse_using(source,caller,parser,for_parser(parser))
}
pub(crate) fn parse_using(source:&str,caller:&Path,parser:ParseFn,preparer:Option<FrontendPreparer>)->Result<PreparedModule,String>{
    let mut emitted=None;
    let mut result=parse_using_with_sink(source,caller,parser,preparer,&mut |hints|emitted=Some(hints))?;
    if result.hints.is_none(){result.hints=emitted;}
    Ok(result)
}
pub(crate) fn parse_using_with_sink(source:&str,caller:&Path,parser:ParseFn,preparer:Option<FrontendPreparer>,sink:&mut PreparationSink<'_>)->Result<PreparedModule,String>{
    // Capture before invoking user hooks/parser. Re-registration during parsing
    // must not splice a newer producer into an older source/context receipt.
    let binding=preparer.map_or_else(||crate::statement_codegen::selected(parser),|p|p.producer);
    if let Some(preparer)=preparer{
        if !preparer.matches_parser(parser){return Err("frontend preparation parser mismatch".into());}
        let mut emitted=false;
        let mut repeated=false;
        let mut forward=|hints| {
            if emitted { repeated=true; } else { emitted=true; sink(hints); }
        };
        let (mut result,context)=if let Some(prepare)=binding.and_then(|b|b.prepare()) {
            let (result,context)=prepare(source,caller,&mut forward)?;(result,Some(context))
        }else{(preparer.invoke(source,caller,&mut forward)?,None)};
        let conflicting=emitted && result.hints.is_some();
        if conflicting { result.hints=None; }
        let complete=!repeated && !conflicting && (emitted||result.hints.is_some());
        Ok(PreparedModule{module:result.module,hints:result.hints,preparer:Some(preparer),preparation_hints_complete:complete,
            statement_production:binding.and_then(|binding|binding.capture(source,context))})
    }else{
        let module=parser(source)?;
        Ok(PreparedModule{module,hints:None,preparer:None,preparation_hints_complete:false,
            statement_production:binding.and_then(|binding|binding.capture(source,None))})
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::Lang;
    fn module(source: &str) -> Module {
        Module { name: source.into(), language: Lang::Unknown, body: vec![], imports: vec![],
            directives: Default::default(), canon: Default::default() }
    }
    struct Hints;
    impl DeferredPreparationHints for Hints {
        fn take_ready_paths(&mut self, _: &mut dyn FnMut(&str) -> Option<String>) -> Vec<String> { vec![] }
        fn is_empty(&self) -> bool { true }
    }
    fn legacy(source: &str) -> Result<Module, String> { Ok(module(source)) }
    fn selected(_: &str) -> Result<Module, String> { panic!("registered provider must parse once") }
    fn committed(source: &str, _: &Path, sink: &mut PreparationSink<'_>) -> Result<FrontendPreparation,String> {
        sink(Box::new(Hints));
        Ok(FrontendPreparation { module: module(source), hints: None })
    }
    fn repeated(source: &str, _: &Path, sink: &mut PreparationSink<'_>) -> Result<FrontendPreparation,String> {
        sink(Box::new(Hints)); sink(Box::new(Hints));
        Ok(FrontendPreparation { module: module(source), hints: Some(Box::new(Hints)) })
    }
    fn failed(_: &str, _: &Path, sink: &mut PreparationSink<'_>) -> Result<FrontendPreparation,String> {
        sink(Box::new(Hints)); Err("normalization failed".into())
    }
    fn resolver(_: &Path, _: &Path, target: &str, _: bool) -> std::path::PathBuf { target.into() }
    fn bootstrap_policy_parser(source:&str)->Result<Module,String>{
        Ok(module(&format!("bootstrap-policy:{source}")))
    }
    #[test]
    fn repeated_bootstrap_preserves_explicit_policy_and_disable() {
        let defaults=Capabilities { prepare:Some(committed), early_hints:true, ..Default::default() };
        register_default(bootstrap_policy_parser,defaults);
        let original=snapshot(bootstrap_policy_parser);
        assert!(original.preparer().is_some());
        // Explicit selection supersedes defaults, even when it is an empty
        // policy deliberately disabling all optimized preparation hooks.
        register(bootstrap_policy_parser,Capabilities::default());
        let disabled=snapshot(bootstrap_policy_parser);
        assert_ne!(original.identity(),disabled.identity());
        register_default(bootstrap_policy_parser,defaults);
        assert_eq!(snapshot(bootstrap_policy_parser).identity(),disabled.identity());
        assert!(snapshot(bootstrap_policy_parser).preparer().is_none());
        // A later explicit replacement still works and leaves prior receipts
        // immutable. Bootstrap cannot overwrite the selected resolver/setup.
        register(bootstrap_policy_parser,Capabilities { resolve_path:Some(resolver),request_setup_reuse:true,..Default::default() });
        let selected=snapshot(bootstrap_policy_parser);
        register_default(bootstrap_policy_parser,defaults);
        assert_eq!(snapshot(bootstrap_policy_parser).identity(),selected.identity());
        assert!(selected.capabilities.request_setup_reuse);
        assert!(disabled.preparer().is_none());
        assert!(original.preparer().is_some());
    }
    #[test]
    fn operand_emitter_registration_is_snapshotted_and_invalidates_identity() {
        fn parser(source: &str) -> Result<Module, String> { Ok(module(&format!("operand-policy:{source}"))) }
        fn producer(_: &str, _: &mut Vec<vybe_runtime::Chunk>, _: usize, _: u8, _: u32) -> bool { false }
        fn specialize(_: &str, _: &mut Vec<vybe_runtime::Chunk>, _: usize, _: u8, _: u32,
            _: vybe_runtime::registry::emitter_results::EmitRequest,
            _: &[crate::primitives::argument_emission::ProducedValueProof])
            -> vybe_runtime::registry::emitter_results::EmitOutcome {
            vybe_runtime::registry::emitter_results::EmitOutcome::Unhandled
        }
        static DEFINITIONS: &[crate::primitives::argument_emission::Definition] = &[
            crate::primitives::argument_emission::Definition {
                operation: "operand_test.operation", parser, producer, dispatch: specialize,
            },
        ];
        register(parser, Capabilities::default());
        let before = snapshot(parser);
        assert!(before.capabilities.argument_emitters.is_empty());
        register(parser, Capabilities { argument_emitters: DEFINITIONS, ..Default::default() });
        let selected = snapshot(parser);
        assert_ne!(before.identity(), selected.identity());
        assert!(before.capabilities.argument_emitters.is_empty());
        assert!(selected.argument_selection().matches_parser(parser));
        assert!(!selected.argument_selection().matches_parser(legacy));
        register(parser, Capabilities::default());
        assert_eq!(snapshot(parser).identity(), before.identity());
        assert_eq!(selected.capabilities.argument_emitters.len(), 1);
    }
    #[test]
    fn no_hook_preserves_legacy_module_and_discovery() {
        let result = parse_using("legacy", Path::new("."), legacy, None).unwrap();
        assert_eq!(result.module.name, "legacy");
        assert!(result.hints.is_none());
        assert!(!result.preparation_hints_complete);
        assert!(result.preparer.is_none());
    }
    #[test]
    fn committed_batch_is_delivered_before_module_return_without_second_parse() {
        let preparer = FrontendPreparer { parser: selected, prepare: committed, early_hints: false, producer:None };
        let mut deliveries = 0;
        let result = parse_using_with_sink("module", Path::new("."), selected, Some(preparer),
            &mut |_| deliveries += 1).unwrap();
        assert_eq!(deliveries, 1);
        assert_eq!(result.module.name, "module");
        assert!(result.preparation_hints_complete);
        assert!(result.hints.is_none());
    }
    #[test]
    fn invalid_protocol_retains_module_and_restores_discovery() {
        let preparer = FrontendPreparer { parser: selected, prepare: repeated, early_hints: false, producer:None };
        let mut deliveries = 0;
        let result = parse_using_with_sink("module", Path::new("."), selected, Some(preparer),
            &mut |_| deliveries += 1).unwrap();
        assert_eq!(deliveries, 1);
        assert_eq!(result.module.name, "module");
        assert!(!result.preparation_hints_complete);
        assert!(result.hints.is_none());
    }
    #[test]
    fn committed_speculation_does_not_turn_normalization_failure_into_success() {
        let preparer = FrontendPreparer { parser: selected, prepare: failed, early_hints: false, producer:None };
        let mut deliveries = 0;
        let result = parse_using_with_sink("bad", Path::new("."), selected, Some(preparer),
            &mut |_| deliveries += 1);
        assert_eq!(deliveries, 1);
        assert_eq!(result.err().as_deref(), Some("normalization failed"));
    }
    #[test]
    fn snapshot_keeps_all_capabilities_from_one_registration() {
        // This unique parser is private to this test, avoiding registry races.
        fn identity_parser(source: &str) -> Result<Module,String> { Ok(module(&format!("identity-policy:{source}"))) }
        register(identity_parser, Capabilities { prepare: Some(committed), ..Default::default() });
        let old = snapshot(identity_parser);
        register(identity_parser, Capabilities { prepare: Some(repeated), resolve_path: Some(resolver),
            request_setup_reuse: true, early_hints: true, ..Default::default() });
        let current = snapshot(identity_parser);
        assert_ne!(old.identity(), current.identity());
        assert!(std::ptr::fn_addr_eq(old.capabilities.prepare.unwrap(), committed as PrepareFn));
        assert!(old.capabilities.resolve_path.is_none());
        assert!(!old.capabilities.request_setup_reuse);
        assert!(!old.capabilities.early_hints);
        assert!(!old.preparer().unwrap().matches(current.preparer().unwrap()));
        assert!(parse_using("module", Path::new("."), legacy, old.preparer()).is_err());
        register(identity_parser, Capabilities::default());
        assert!(for_parser(identity_parser).is_none());
    }
    fn context_prepare(source:&str,caller:&Path,_:&mut PreparationSink<'_>)->Result<(FrontendPreparation,crate::statement_codegen::ProducerContext),String>{
        Ok((FrontendPreparation{module:module(source),hints:None},std::sync::Arc::new((source.to_owned(),caller.to_path_buf()))))
    }
    fn context_produce(source:&str,module:Module,context:&crate::statement_codegen::ProducerContext,_:&mut crate::statement_codegen::StatementSink<'_>)->Result<(),String>{
        let (original,caller)=context.downcast_ref::<(String,std::path::PathBuf)>().ok_or("wrong context type")?;
        if original!=source || module.name!=source || caller!=Path::new("caller.php"){return Err("source/context mismatch".into());}Ok(())
    }
    static CONTEXT_HOOKS:crate::statement_codegen::PreparedHooks=crate::statement_codegen::PreparedHooks{prepare:context_prepare,produce:context_produce};
    fn forbidden_prepare(_: &str,_:&Path,_:&mut PreparationSink<'_>)->Result<FrontendPreparation,String>{panic!("legacy prepare must not run when contextual hooks selected")}
    #[test]
    fn prepared_producer_retains_request_context_and_exact_source() {
        fn parser(source:&str)->Result<Module,String>{Ok(module(source))}
        register(parser,Capabilities{prepare:Some(forbidden_prepare),..Default::default()});
        crate::statement_codegen::register_prepared(parser,&CONTEXT_HOOKS);
        let registration=snapshot(parser);let id=registration.identity();
        crate::statement_codegen::register_prepared(parser,&CONTEXT_HOOKS);
        assert_eq!(snapshot(parser).identity(),id,"identical bootstrap must not invalidate source receipts");
        let result=parse_using("source A",Path::new("caller.php"),parser,registration.preparer()).unwrap();
        let production=result.statement_production.unwrap();let cloned=production.clone();
        production.emit(result.module,&mut |_,_|Ok(())).unwrap();
        assert!(cloned.emit(module("source B"),&mut |_,_|Ok(())).unwrap_err().contains("source/context mismatch"));
        let independent=parse_using("source B",Path::new("caller.php"),parser,registration.preparer()).unwrap();
        independent.statement_production.unwrap().emit(independent.module,&mut |_,_|Ok(())).unwrap();
    }
    #[test]
    fn disabled_preparer_does_not_invoke_context_hooks() {
        fn parser(source:&str)->Result<Module,String>{Ok(module(source))}
        register(parser,Capabilities::default());crate::statement_codegen::register_prepared(parser,&CONTEXT_HOOKS);
        let result=parse_using("legacy",Path::new("."),parser,for_parser(parser)).unwrap();
        assert_eq!(result.module.name,"legacy");assert!(result.preparer.is_none());assert!(result.statement_production.is_none());
    }
    #[test]
    fn legacy_producer_binding_is_captured_before_parser_reregistration() {
        fn first(_: &str,_:Module,_:&mut crate::statement_codegen::StatementSink<'_>)->Result<(),String>{Ok(())}
        fn replacement(_: &str,_:Module,_:&mut crate::statement_codegen::StatementSink<'_>)->Result<(),String>{Err("new binding".into())}
        fn parser(source:&str)->Result<Module,String>{crate::statement_codegen::register(parser,replacement);Ok(module(source))}
        crate::statement_codegen::register(parser,first);
        let old=snapshot(parser);let result=parse_using("source",Path::new("."),parser,None).unwrap();
        assert_ne!(old.identity(),snapshot(parser).identity());assert!(old.identity()[9]>0);
        result.statement_production.unwrap().emit(result.module,&mut |_,_|Ok(())).unwrap();
    }
    #[test]
    fn prepared_binding_and_identity_survive_registry_change_during_prepare() {
        fn parser(source:&str)->Result<Module,String>{Ok(module(source))}
        fn replacement(_: &str,_:Module,_:&mut crate::statement_codegen::StatementSink<'_>)->Result<(),String>{Err("replacement must not consume old context".into())}
        fn prepare(source:&str,caller:&Path,sink:&mut PreparationSink<'_>)->Result<(FrontendPreparation,crate::statement_codegen::ProducerContext),String>{
            crate::statement_codegen::register(parser,replacement);context_prepare(source,caller,sink)
        }
        static HOOKS:crate::statement_codegen::PreparedHooks=crate::statement_codegen::PreparedHooks{prepare,produce:context_produce};
        register(parser,Capabilities{prepare:Some(forbidden_prepare),..Default::default()});
        crate::statement_codegen::register_prepared(parser,&HOOKS);
        let old=snapshot(parser);let oldid=old.identity();
        let result=parse_using("source",Path::new("caller.php"),parser,old.preparer()).unwrap();
        let current=snapshot(parser);assert_ne!(oldid,current.identity());assert_eq!(oldid,old.identity());
        assert!(!old.preparer().unwrap().matches(current.preparer().unwrap()));
        result.statement_production.unwrap().emit(result.module,&mut |_,_|Ok(())).unwrap();
    }

}
