//! Single-parse composition with an existing journaled AST builder.
use super::{Binding,Collector};
use crate::{builder::{Builder,BuildError},grammar::RuleId,Span};
use std::collections::HashSet;

struct Undo { ast:usize, collector:usize }
/// Add preparation-intent hooks to an AST builder without another parse or
/// capture/AST walk. The AST builder must use an additive monotonic journal
/// cursor; checkpoint calls stay pure and allocate nothing.
///
/// With bound intents, shortcuts that may omit child hooks require explicit
/// admission. Register only Pratt/atomic rules whose omitted callbacks cannot
/// hide an intent's selector or operand. Other regions retain PEG matching.
pub struct WithBuilder<'a,B> {
    pub ast:B,
    pub collector:Collector<'a>,
    bound:bool,
    undo:Vec<Undo>,
    pratt:HashSet<RuleId>,
    opaque:HashSet<RuleId>,
    pratt_admission:Option<fn(&B,RuleId,usize,&str)->bool>,
}
impl<'a,B:Builder> WithBuilder<'a,B> {
    pub fn new(ast:B,bindings:impl IntoIterator<Item=Binding<'a>>)->Self {
        let bindings:Vec<_>=bindings.into_iter().collect();
        Self {ast,bound:!bindings.is_empty(),collector:Collector::new(bindings),undo:Vec::new(),pratt:HashSet::new(),opaque:HashSet::new(),pratt_admission:None}
    }
    pub fn allow_pratt_rules(&mut self,rules:impl IntoIterator<Item=RuleId>) {self.pratt.extend(rules);}
    pub fn allow_opaque_rules(&mut self,rules:impl IntoIterator<Item=RuleId>) {self.opaque.extend(rules);}
    /// Pure source admission for islands whose omitted hooks and contextual
    /// recapture cannot hide an intent. The AST owner's policy still applies.
    pub fn set_pratt_admission(&mut self,admission:fn(&B,RuleId,usize,&str)->bool) {self.pratt_admission=Some(admission);}
    fn collect(&mut self,operation:impl FnOnce(&mut Collector<'a>)->Result<(),BuildError>)->Result<(),BuildError> {
        let ast=self.ast.checkpoint();let collector=self.collector.checkpoint();
        let result=operation(&mut self.collector);
        if self.collector.checkpoint()!=collector {self.undo.push(Undo {ast,collector});}
        result
    }
}
impl<B:Builder> Builder for WithBuilder<'_,B> {
    fn checkpoint(&self)->usize {self.ast.checkpoint()+self.undo.len()}
    fn rollback(&mut self,mark:usize) {
        while self.undo.last().is_some_and(|undo|undo.ast+self.undo.len()-1>=mark) {
            let undo=self.undo.pop().unwrap();
            self.ast.rollback(undo.ast);self.collector.rollback(undo.collector);
        }
        self.ast.rollback(mark-self.undo.len());
    }
    fn observes_rule(&self,rule:RuleId)->bool {self.ast.observes_rule(rule)||self.collector.observes_rule(rule)}
    fn observes_capture(&self,rule:RuleId)->bool {self.ast.observes_capture(rule)}
    fn try_begin(&mut self,rule:RuleId,offset:usize)->Result<(),BuildError> {
        if self.ast.observes_rule(rule) {self.ast.try_begin(rule,offset)?;}
        if !self.collector.observes_rule(rule) {return Ok(());}
        self.collect(|collector|collector.try_begin(rule,offset))
    }
    fn try_finish(&mut self,rule:RuleId,span:Span,source:&str)->Result<(),BuildError> {
        if self.ast.observes_rule(rule) {self.ast.try_finish(rule,span,source)?;}
        if !self.collector.observes_rule(rule) {return Ok(());}
        self.collect(|collector|collector.try_finish(rule,span,source))
    }
    fn try_finish_capture(&mut self,rule:RuleId,span:Span,source:&str,capture:Option<crate::tree::CaptureHandle>)->Result<(),BuildError> {
        if self.ast.observes_rule(rule) {self.ast.try_finish_capture(rule,span,source,capture)?;}
        if !self.collector.observes_rule(rule) {return Ok(());}
        self.collect(|collector|collector.try_finish_capture(rule,span,source,capture))
    }
    fn begin(&mut self,rule:RuleId,offset:usize) {
        if !self.bound {self.ast.begin(rule,offset);return;}
        self.try_begin(rule,offset).expect("checked source-unit begin hook");
    }
    fn finish(&mut self,rule:RuleId,span:Span,source:&str) {
        if !self.bound {self.ast.finish(rule,span,source);return;}
        self.try_finish(rule,span,source).expect("checked source-unit finish hook");
    }
    fn committed(&mut self,source:&str,consumed:usize) {self.ast.committed(source,consumed);self.collector.committed(source,consumed);self.undo.clear();}
    fn supports_choice_heads(&self)->bool {self.ast.supports_choice_heads()}
    fn supports_unobserved_failure_prefixes(&self)->bool {self.ast.supports_unobserved_failure_prefixes()}
    fn failure_prefix_hooks_are_inert(&self,rule:RuleId)->bool {!self.collector.observes_rule(rule)&&self.ast.failure_prefix_hooks_are_inert(rule)}
    fn supports_shared_prefix(&self,rule:RuleId)->bool {!self.bound&&self.ast.supports_shared_prefix(rule)}
    fn supports_shared_inner_prefix(&self,r:RuleId)->bool {!self.bound&&self.ast.supports_shared_inner_prefix(r)}
    fn supports_shared_group_prefix(&self,a:RuleId,b:RuleId)->bool {!self.bound&&self.ast.supports_shared_group_prefix(a,b)}
    fn supports_shared_wrapper_prefix(&self,a:RuleId,b:RuleId)->bool {!self.bound&&self.ast.supports_shared_wrapper_prefix(a,b)}
    fn supports_shared_wrapper_prefix_at(&self,a:RuleId,b:RuleId,offset:usize,source:&str)->bool {!self.bound&&self.ast.supports_shared_wrapper_prefix_at(a,b,offset,source)}
    fn supports_pratt(&self,rule:RuleId)->bool {(!self.bound||self.pratt.contains(&rule))&&self.ast.supports_pratt(rule)}
    fn supports_pratt_at(&self,rule:RuleId,offset:usize,source:&str)->bool {(!self.bound||self.pratt.contains(&rule)||self.pratt_admission.is_some_and(|admit|admit(&self.ast,rule,offset,source)))&&self.ast.supports_pratt_at(rule,offset,source)}
    fn accepts_pratt(&self,rule:RuleId,span:Span)->bool {self.ast.accepts_pratt(rule,span)}
    fn suppress_pratt_captures(&self,rule:RuleId)->bool {self.ast.suppress_pratt_captures(rule)}
    fn opaque_rule(&self,rule:RuleId)->bool {(!self.bound||self.opaque.contains(&rule))&&self.ast.opaque_rule(rule)}
    fn prune_capture_children(&self,rule:RuleId,span:Span)->bool {self.ast.prune_capture_children(rule,span)}
    fn try_reduce(&mut self,rule:RuleId,fixity:crate::pratt::Fixity,span:Span)->Result<(),BuildError> {self.ast.try_reduce(rule,fixity,span)}
    fn reduce(&mut self,rule:RuleId,fixity:crate::pratt::Fixity,span:Span) {self.ast.reduce(rule,fixity,span);}
    fn try_pratt_boundary(&mut self,rule:RuleId,offset:usize)->Result<(),BuildError> {self.ast.try_pratt_boundary(rule,offset)}
}
