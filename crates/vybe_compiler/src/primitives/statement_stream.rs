//! Source-unit instruction emission driven by a validated typed producer.
//! Preparation discovers forward semantic facts once. The stream retains no
//! Module/body borrow; callbacks transfer statements and release them after emit.
use super::*;
use std::hash::{Hash, Hasher};

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StatementProgress {
    pub source_index: usize,
    pub emitted: bool,
    pub appended_chunks: std::ops::Range<usize>,
    pub entry_bytes: std::ops::Range<usize>,
}

// A bounded receipt, not a serialized AST or persistent/security certificate.
// The registered producer/source identity is validated by its caller. Including
// the complete normalized shape/spans catches accidental mismatched producers.
#[derive(Clone, Copy, PartialEq, Eq)]
struct Shape([u64; 2]);
fn shape(stmt: &Statement) -> Shape {
    struct Digest(std::hash::DefaultHasher, std::hash::DefaultHasher);
    impl std::fmt::Write for Digest {
        fn write_str(&mut self, text: &str) -> std::fmt::Result {
            self.0.write(text.as_bytes()); self.1.write(text.as_bytes()); Ok(())
        }
    }
    let mut digest=Digest(std::hash::DefaultHasher::new(),std::hash::DefaultHasher::new());
    "statement-stream-0".hash(&mut digest.0);"statement-stream-1".hash(&mut digest.1);
    std::fmt::write(&mut digest,format_args!("{stmt:?}")).expect("infallible statement digest");
    Shape([digest.0.finish(),digest.1.finish()])
}
enum Action { Emit, HoistedFunction, MergedClass(Box<Statement>), MergedDuplicate }
struct Expected { shape: Option<Shape>, action: Action }

/// One validated SOURCE UNIT, not a whole interactive request/program.
/// Hoisted functions have already been emitted using the ordinary Compiler.
/// Every original manifest statement must subsequently arrive once, in order;
/// function callbacks validate their receipt but do not emit a second body.
/// Pending chunks remain unrelocated/unpublished until successful `finish`.
pub struct StatementEmission {
    compiler: Option<Compiler>,
    expected: Vec<Option<Expected>>,
    next: usize,
    builtin_signatures: Option<Statement>,
    component_source: bool,
    component_binary: Option<std::sync::Arc<vybe_ast::canon::ComponentBinary>>,
    timing: Option<String>,
    total_started: Instant,
    body_elapsed: Duration,
    module_producers: Option<std::collections::BTreeMap<&'static str,(usize,Duration)>>,
}
/// Private owned capability: callers can advance or abort this session, but
/// cannot obtain, mutate or replace its original statement iterator/inner sink.
/// Preparation/actions are identical to the checked independent producer path.
pub(crate) struct OwnedStatementEmission {
    emission: StatementEmission,
    statements: std::vec::IntoIter<Statement>,
}
impl OwnedStatementEmission {
    pub(crate) fn emit_next(&mut self)->Result<Option<StatementProgress>,String>{
        if self.emission.compiler.is_none(){return Err("statement stream is poisoned".into());}
        let Some(statement)=self.statements.next()else{return Ok(None)};
        let index=self.emission.next;
        let result=self.emission.emit_owned_inner(index,&statement);
        if result.is_err(){self.emission.compiler.take();}
        result.map(Some)
    }
    pub(crate) fn finish(self)->Result<CompileResult,String>{
        if self.statements.len()!=0{return Err("owned statement stream has pending original statements".into());}
        self.emission.finish()
    }
    #[allow(dead_code)]
    pub(crate) fn abort(self){}
}

/// Compiler-private immutable view of the exact prepared manifest. The borrow
/// prevents source refinement/restoration until emission has finished or aborts.
/// No independent frontend callback can replace or mutate these iterator nodes.
pub(crate) struct BorrowedManifestEmission<'a> {
    emission: StatementEmission,
    statements: std::slice::Iter<'a,Statement>,
}
impl BorrowedManifestEmission<'_> {
    pub(crate) fn emit_next(&mut self)->Result<Option<StatementProgress>,String>{
        if self.emission.compiler.is_none(){return Err("statement stream is poisoned".into());}
        let Some(statement)=self.statements.next()else{return Ok(None)};
        let index=self.emission.next;
        let result=self.emission.emit_owned_inner(index,statement);
        if result.is_err(){self.emission.compiler.take();}
        result.map(Some)
    }
    pub(crate) fn finish(self)->Result<CompileResult,String>{
        if self.statements.len()!=0{return Err("borrowed statement stream has pending original statements".into());}
        self.emission.finish()
    }
    #[allow(dead_code)]pub(crate) fn abort(self){}
}

impl Compiler {
    /// Discover semantic/header facts once, then accept independently delivered
    /// normalized statements. No source-unit AST reference survives this call.
    pub fn begin_statement_emission(self, module:&Module)->Result<StatementEmission,String>{
        let prepared=self.begin_emission(module)?;
        StatementEmission::prepare(prepared,module)
    }
    /// Compiler-only ownership transfer. Borrow the SAME semantic preparation
    /// route while sealing actions, then move the untouched original body into
    /// the private iterator. No independent frontend producer can replace it.
    pub(crate) fn begin_owned_statement_emission_prepared(self,module:Module,
        parser:vybe_runtime::registry::ParseFn)->Result<OwnedStatementEmission,String>{
        let emission={
            let prepared=self.begin_emission_prepared(&module,parser)?;
            StatementEmission::prepare_with_origin(prepared,&module,false)?
        };
        Ok(OwnedStatementEmission{emission,statements:module.body.into_iter()})
    }
    /// Immutable compiler-owned source view; source identity is validated by
    /// Bundle before entry, using the same normalizer/hoist/action setup.
    pub(crate) fn begin_borrowed_manifest_emission_prepared<'a>(self,module:&'a Module,
        parser:vybe_runtime::registry::ParseFn)->Result<BorrowedManifestEmission<'a>,String>{
        let prepared=self.begin_emission_prepared(module,parser)?;
        let emission=StatementEmission::prepare_with_origin(prepared,module,false)?;
        Ok(BorrowedManifestEmission{emission,statements:module.body.iter()})
    }
    /// Registered CaaS entry; preserve the same parser/normalizer admission as
    /// ordinary borrowed prepared compilation, without donating source bodies.
    pub(crate) fn begin_statement_emission_prepared(self,module:&Module,
        parser:vybe_runtime::registry::ParseFn)->Result<StatementEmission,String>{
        let prepared=self.begin_emission_prepared(module,parser)?;
        StatementEmission::prepare(prepared,module)
    }
}
impl StatementEmission {
    fn prepare(prepared:PreparedEmission<'_>,module:&Module)->Result<Self,String>{
        Self::prepare_with_origin(prepared,module,true)
    }
    fn prepare_with_origin(prepared:PreparedEmission<'_>,module:&Module,independent:bool)->Result<Self,String>{
        let prepared=prepared.into_statement_stream_parts()?;
        let compiler=prepared.compiler;
        let has_partial=module.body.iter().any(|s|matches!(&s.kind,StmtKind::ClassDecl{modifiers,..}if modifiers.is_partial));
        let key=|name:&str|if compiler.case_sensitive{name.to_owned()}else{name.to_lowercase()};
        let mut merged=std::collections::HashMap::new();
        if has_partial {
            // Existing preparation already performed exactly the legacy merge.
            // Retain only its grouped class payloads, not the merged Module.
            for statement in prepared.body.into_owned(){
                if let StmtKind::ClassDecl{name,..}=&statement.kind{merged.insert(key(name),statement);}
            }
        }
        let mut seen=std::collections::HashSet::new();
        let mut expected=Vec::with_capacity(module.body.len());
        for statement in &module.body{
            let action=if matches!(statement.kind,StmtKind::FunctionDecl{..}){Action::HoistedFunction}
                else if has_partial {if let StmtKind::ClassDecl{name,..}=&statement.kind{
                    let name=key(name);
                    if seen.insert(name.clone()){Action::MergedClass(Box::new(merged.remove(&name).ok_or("missing partial class group")?))}
                    else{Action::MergedDuplicate}
                }else{Action::Emit}}else{Action::Emit};
            expected.push(Some(Expected{shape:independent.then(||shape(statement)),action}));
        }
        Ok(Self{compiler:Some(compiler),expected,next:0,
            builtin_signatures:prepared.builtin_signatures,
            component_source:prepared.component_source,component_binary:prepared.component_binary,
            module_producers:prepared.timing.as_ref().map(|_|std::collections::BTreeMap::new()),
            timing:prepared.timing,total_started:prepared.total_started,body_elapsed:Duration::ZERO})
    }
    /// Inspection only. Import/global/type operands are still unit-relative;
    /// callers must not install or execute this provisional instruction stream.
    pub fn pending_chunks(&self)->Result<&[Chunk],String>{
        self.compiler.as_ref().map(|c|c.chunks.as_slice()).ok_or_else(||"statement stream is poisoned".into())
    }
    pub fn expected_statement_count(&self)->usize{self.expected.len()}
    pub fn consumed_statement_count(&self)->usize{self.next}
    /// Consume one normalized typed statement from the registered producer.
    /// All Stmt/Expr kinds use the existing emitter; this never compiles a
    /// one-statement Module or runs the whole semantic setup again.
    pub fn emit_statement(&mut self,index:usize,statement:Statement)->Result<StatementProgress,String>{
        let result=self.emit_inner(index,statement);
        if result.is_err(){self.compiler.take();}
        result
    }
    fn emit_inner(&mut self,index:usize,statement:Statement)->Result<StatementProgress,String>{
        if self.compiler.is_none(){return Err("statement stream is poisoned".into());}
        if index!=self.next{return Err(format!("statement producer out of order: expected {}, received {index}",self.next));}
        let row=self.expected.get_mut(index).and_then(Option::take).ok_or_else(||format!("unexpected statement producer {index}"))?;
        let expected=row.shape.ok_or("independent callback cannot consume compiler-owned statement slots")?;
        if expected!=shape(&statement){return Err(format!("statement producer differs from sealed normalized manifest at {index}"));}
        self.emit_action(row.action,&statement,index)
    }
    /// Only compiler-private original-manifest iterators reach this method.
    /// The node is borrowed from the exact vector sealed in preparation.
    fn emit_owned_inner(&mut self,index:usize,statement:&Statement)->Result<StatementProgress,String>{
        if self.compiler.is_none(){return Err("statement stream is poisoned".into());}
        if index!=self.next{return Err("owned statement order mismatch".into());}
        let row=self.expected.get_mut(index).and_then(Option::take).ok_or("unexpected owned statement")?;
        if row.shape.is_some(){return Err("owned iterator cannot consume independent callback slots".into());}
        self.emit_action(row.action,statement,index)
    }
    fn emit_action(&mut self,action:Action,statement:&Statement,index:usize)->Result<StatementProgress,String>{
        let compiler=self.compiler.as_ref().unwrap();
        let before_chunks=compiler.chunks.len();let before_bytes=compiler.chunks[0].code.len();
        let emitted=match action{
            Action::Emit=>{self.emit_body(statement)?;true},
            Action::MergedClass(merged)=>{self.emit_body(&merged)?;true},
            Action::HoistedFunction|Action::MergedDuplicate=>false,
        };
        self.next+=1;
        let compiler=self.compiler.as_ref().unwrap();
        Ok(StatementProgress{source_index:index,emitted,appended_chunks:before_chunks..compiler.chunks.len(),entry_bytes:before_bytes..compiler.chunks[0].code.len()})
    }
    fn emit_body(&mut self,statement:&Statement)->Result<(),String>{
        let compiler=self.compiler.as_mut().ok_or("statement stream is poisoned")?;
        let _lines=vybe_runtime::source_lines::LineEmissionScope::enter(compiler.compact_source_lines);
        let _total=vybe_runtime::debugger::DebugPhase::current("compiler total");
        let _phase=vybe_runtime::debugger::DebugPhase::current("compiler module body");
        let started=self.timing.as_ref().map(|_|Instant::now());
        compiler.compile_stmt(statement)?;
        if let (Some(path),Some(started))=(self.timing.as_deref(),started){
            let elapsed=started.elapsed();self.body_elapsed+=elapsed;
            let kind=match &statement.kind{StmtKind::ClassDecl{..}=>"class",StmtKind::Expr(_)=>"expression",StmtKind::VarDecl{..}=>"variable",StmtKind::If{..}=>"conditional",StmtKind::Block(_)=>"block",StmtKind::Try{..}=>"try",_=>"other"};
            let aggregate=self.module_producers.as_mut().unwrap().entry(kind).or_default();aggregate.0+=1;aggregate.1+=elapsed;
            if let StmtKind::ClassDecl{name,..}=&statement.kind{write_compiler_timing(path,&format!("compile_class_decl {name} line={}",statement.span.start_line),elapsed);}
        }
        Ok(())
    }
    /// Called only after the selected grammar entry succeeds and every producer
    /// is received. Resolves native instruction fixups once; no Wasm export.
    pub fn finish_with(mut self, provider: impl FnOnce(
        &mut crate::statement_codegen::StatementSink<'_>,
    ) -> Result<(), String>) -> Result<CompileResult, String> {
        provider(&mut |index, statement| self.emit_statement(index, statement).map(|_| ()))?;
        self.finish()
    }

    /// Finalize a fully delivered, successfully parsed source unit.
    pub fn finish(mut self)->Result<CompileResult,String>{
        if self.compiler.is_none(){return Err("statement stream is poisoned".into());}
        if self.next!=self.expected.len(){return Err(format!("statement producer incomplete: received {} of {}",self.next,self.expected.len()));}
        if let Some(statement)=self.builtin_signatures.take(){self.emit_body(&statement)?;}
        if let Some(path)=self.timing.as_deref(){
            for(kind,(count,elapsed))in self.module_producers.take().unwrap(){write_compiler_timing(path,&format!("compile_module_producer_kind {kind} count={count}"),elapsed);}
            write_compiler_timing(path,"compile_non_function_stmts",self.body_elapsed);
        }
        self.compiler.take().unwrap().finish_emission(self.component_source,self.component_binary,self.timing,self.total_started)
    }
    /// Abandon all unpublished code after parser failure. Previously installed
    /// source units belong to the caller/VM and are not touched by this stream.
    pub fn abort(self){}
}

#[cfg(test)]
mod owned_manifest_tests {
    use super::*;
    fn module() -> Module {
        Module { name:"owned-test".into(),language:vybe_ast::Lang::Unknown,
            body:vec![Statement::new(StmtKind::Empty),Statement::new(StmtKind::Empty)],
            imports:vec![],directives:Default::default(),canon:Default::default() }
    }
    fn parser(_: &str)->Result<Module,String>{Ok(module())}
    fn compiler()->Compiler {
        Compiler::with_profile(crate::profile::parse_profile("[info]\nname='owned-test'\n[compiler]\n").unwrap())
            .with_run_module_end_hooks(false)
    }
    #[test]
    fn owned_original_iterator_matches_checked_emission_and_order() {
        let original=module();let checked_module=original.clone();
        let checked=compiler().begin_statement_emission(&checked_module).unwrap()
            .finish_with(|sink|{for(index,statement)in checked_module.body.clone().into_iter().enumerate(){sink(index,statement)?;}Ok(())}).unwrap();
        let mut owned=compiler().begin_owned_statement_emission_prepared(original,parser).unwrap();
        assert_eq!(owned.emit_next().unwrap().unwrap().source_index,0);
        assert_eq!(owned.emit_next().unwrap().unwrap().source_index,1);
        assert!(owned.emit_next().unwrap().is_none());
        let actual=owned.finish().unwrap();assert_eq!(checked.chunks.len(),actual.chunks.len());
        for(a,b)in checked.chunks.iter().zip(&actual.chunks){assert_eq!(a.code,b.code);assert_eq!(a.lines,b.lines);}
    }
    #[test]
    fn public_independent_callback_cannot_use_owned_slots() {
        let mut owned=compiler().begin_owned_statement_emission_prepared(module(),parser).unwrap();
        let error=owned.emission.emit_statement(0,Statement::new(StmtKind::Empty)).unwrap_err();
        assert!(error.contains("independent callback"));
        assert!(owned.emit_next().unwrap_err().contains("poisoned"));
    }
    #[test]
    fn independent_producer_still_rejects_changed_and_reordered_nodes() {
        let source=module();let mut changed=compiler().begin_statement_emission(&source).unwrap();
        assert!(changed.emit_statement(0,Statement::new(StmtKind::Block(vec![]))).unwrap_err().contains("differs"));
        let mut reordered=compiler().begin_statement_emission(&source).unwrap();
        assert!(reordered.emit_statement(1,Statement::new(StmtKind::Empty)).unwrap_err().contains("out of order"));
        assert!(reordered.pending_chunks().is_err());
    }
    #[test]
    fn owned_finish_refuses_missing_original_nodes_and_abort_discards_session() {
        let owned=compiler().begin_owned_statement_emission_prepared(module(),parser).unwrap();
        assert!(owned.finish().unwrap_err().contains("pending original"));
        let mut owned=compiler().begin_owned_statement_emission_prepared(module(),parser).unwrap();
        owned.emit_next().unwrap();owned.abort();
        // Independent session remains usable; aborted chunks are never returned.
        let mut replacement=compiler().begin_owned_statement_emission_prepared(module(),parser).unwrap();
        while replacement.emit_next().unwrap().is_some(){}assert!(replacement.finish().is_ok());
    }
    #[test]
    fn borrowed_exact_original_body_matches_owned_and_never_transfers_nodes(){
        let source=module();let before=format!("{source:?}");let pointer=source.body.as_ptr();
        let mut borrowed=compiler().begin_borrowed_manifest_emission_prepared(&source,parser).unwrap();
        assert_eq!(borrowed.statements.as_slice().as_ptr(),pointer);
        while borrowed.emit_next().unwrap().is_some(){}let actual=borrowed.finish().unwrap();
        assert_eq!(source.body.as_ptr(),pointer);assert_eq!(format!("{source:?}"),before);
        let mut owned=compiler().begin_owned_statement_emission_prepared(source.clone(),parser).unwrap();
        while owned.emit_next().unwrap().is_some(){}let expected=owned.finish().unwrap();
        assert_eq!(format!("{:?}",actual.chunks),format!("{:?}",expected.chunks));
    }
    #[test]
    fn borrowed_rejects_external_callback_and_missing_nodes_without_mutating_source(){
        let source=module();let before=format!("{source:?}");
        let mut borrowed=compiler().begin_borrowed_manifest_emission_prepared(&source,parser).unwrap();
        assert!(borrowed.emission.emit_statement(0,Statement::new(StmtKind::Empty)).unwrap_err().contains("independent callback"));
        assert!(borrowed.emit_next().unwrap_err().contains("poisoned"));
        let pending=compiler().begin_borrowed_manifest_emission_prepared(&source,parser).unwrap();
        assert!(pending.finish().unwrap_err().contains("pending original"));
        let mut aborted=compiler().begin_borrowed_manifest_emission_prepared(&source,parser).unwrap();aborted.emit_next().unwrap();aborted.abort();
        assert_eq!(format!("{source:?}"),before);
    }

}
