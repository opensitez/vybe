//! A standalone semantic hook consumer for committed preparation intents.
//! Targets are source spans, not evaluated filenames. Embeddings may compose
//! these hooks with their AST builder and replace spans with owned AST handles.
use super::{Binding, Request, PreparationContext};
use crate::{Span, grammar::RuleId, builder::{Builder, BuildError}};
use std::collections::{HashMap,HashSet};

struct Frame { rule: RuleId, start: usize, selector: Option<RuleId>, targets: Vec<(RuleId,Span)> }
enum Undo { PushFrame, PopFrame(Frame), Selector(Option<RuleId>), Target, Pending, Depth(usize) }

/// Explicit opt-in hook collector. Requests are visible only after successful
/// commitment; failed alternatives, predicates and failed parses publish none.
/// Construction and collection never access files, run code, or change scope.
pub struct Collector<'a> {
    bindings: Vec<Binding<'a>>,
    roots: HashMap<RuleId,Vec<usize>>,
    observed: Vec<RuleId>,
    deferred_rules: HashSet<RuleId>,
    deferred_depth: usize,
    frames: Vec<Frame>,
    undo: Vec<Undo>,
    pending: Vec<Request<Span>>,
    committed: Vec<Request<Span>>,
    owned: Vec<Request<String>>,
}
impl<'a> Collector<'a> {
    pub fn new(bindings: impl IntoIterator<Item=Binding<'a>>) -> Self {
        let bindings: Vec<_> = bindings.into_iter().collect();
        let mut roots: HashMap<_,Vec<_>> = HashMap::new();
        let mut observed = Vec::new();
        for (index,binding) in bindings.iter().enumerate() {
            roots.entry(binding.rule).or_default().push(index);
            observed.push(binding.rule);
            observed.push(binding.target);
            observed.extend(binding.selector);
        }
        observed.sort_unstable();observed.dedup();
        Self {bindings,roots,observed,deferred_rules:HashSet::new(),deferred_depth:0,frames:Vec::new(),undo:Vec::new(),pending:Vec::new(),committed:Vec::new(),owned:Vec::new()}
    }
    /// Register frontend-defined deferred body boundaries before matching.
    /// These hooks only classify speculation; they never alter lexical scope,
    /// suppress requests, or decide whether a load executes.
    pub fn defer_within(&mut self,rules:impl IntoIterator<Item=RuleId>) {
        assert!(self.frames.is_empty() && self.undo.is_empty(),"deferred rules must be registered outside a parse");
        for rule in rules {self.deferred_rules.insert(rule);self.observed.push(rule);}
        self.observed.sort_unstable();self.observed.dedup();
    }
    /// Successful finish order, including nested loads. This is preparation
    /// intent only; execution order remains the embedding's AST semantics.
    pub fn requests(&self) -> &[Request<Span>] { &self.committed }
    pub fn take_requests(&mut self) -> Vec<Request<Span>> { std::mem::take(&mut self.committed) }
    /// Unevaluated operand text from the exact committed parser input. This
    /// remains correct when an embedding normalizes/expands source before
    /// parsing; preparation must not index the raw file using normalized spans.
    pub fn owned_requests(&self) -> &[Request<String>] { &self.owned }
    pub fn take_owned_requests(&mut self) -> Vec<Request<String>> { std::mem::take(&mut self.owned) }
    fn record_child(&mut self, rule: RuleId, span: Span) -> Result<(),BuildError> {
        if let Some(frame)=self.frames.last_mut() {
            let indices=&self.roots[&frame.rule];
            if indices.iter().any(|index|self.bindings[*index].selector==Some(rule)) {
                if frame.selector.is_some() {return Err(BuildError {span,message:"ambiguous source-unit selector".into()});}
                self.undo.push(Undo::Selector(frame.selector));
                frame.selector=Some(rule);
            }
            if indices.iter().any(|index|self.bindings[*index].target==rule) {
                frame.targets.push((rule,span));
                self.undo.push(Undo::Target);
            }
        }
        Ok(())
    }
}
impl Builder for Collector<'_> {
    fn checkpoint(&self) -> usize { self.undo.len() }
    fn rollback(&mut self, mark: usize) {
        while self.undo.len()>mark {
            match self.undo.pop().unwrap() {
                Undo::PushFrame => { self.frames.pop().unwrap(); },
                Undo::PopFrame(frame) => self.frames.push(frame),
                Undo::Selector(previous) => self.frames.last_mut().unwrap().selector=previous,
                Undo::Target => {self.frames.last_mut().unwrap().targets.pop().unwrap();},
                Undo::Pending => {self.pending.pop().unwrap();},
                Undo::Depth(previous) => self.deferred_depth=previous,
            }
        }
    }
    fn observes_rule(&self, rule: RuleId) -> bool { self.observed.binary_search(&rule).is_ok() }
    fn begin(&mut self, rule: RuleId, start: usize) {
        if self.deferred_rules.contains(&rule) {
            self.undo.push(Undo::Depth(self.deferred_depth));self.deferred_depth+=1;
        }
        if self.roots.contains_key(&rule) {
            self.frames.push(Frame {rule,start,selector:None,targets:Vec::new()});
            self.undo.push(Undo::PushFrame);
        }
    }
    fn finish(&mut self, _:RuleId, _:Span, _: &str) {}
    fn try_finish(&mut self, rule: RuleId, span: Span, _: &str) -> Result<(),BuildError> {
        let fail = |message: &str| BuildError {span,message:message.into()};
        if self.roots.contains_key(&rule) {
            let frame = self.frames.pop().ok_or_else(||fail("missing source-unit frame"))?;
            let indices = &self.roots[&frame.rule];
            let binding = indices.iter().map(|index|self.bindings[*index])
                .find(|binding|binding.selector==frame.selector);
            let target = binding.and_then(|binding| {
                let mut targets = frame.targets.iter().filter(|(rule,_)|*rule==binding.target);
                let target=targets.next().map(|(_,span)|*span)?;
                targets.next().is_none().then_some(target)
            });
            let valid_frame=frame.rule==rule && frame.start==span.start;
            self.undo.push(Undo::PopFrame(frame));
            if !valid_frame {return Err(fail("mismatched source-unit frame"));}
            let binding=binding.ok_or_else(||fail("missing source-unit selector binding"))?;
            let target=target.ok_or_else(||fail("source-unit requires one semantic target"))?;
            let mut request=binding.request(target,span);
            request.preparation=if self.deferred_depth==0 {PreparationContext::Immediate} else {PreparationContext::Deferred};
            self.pending.push(request);
            self.undo.push(Undo::Pending);
            // A nested load may itself be the parent's declared operand.
            self.record_child(rule,span)?;
        } else {
            self.record_child(rule,span)?;
        }
        if self.deferred_rules.contains(&rule) {
            self.undo.push(Undo::Depth(self.deferred_depth));self.deferred_depth-=1;
        }
        Ok(())
    }
    fn committed(&mut self, source: &str, _: usize) {
        debug_assert!(self.frames.is_empty());
        debug_assert_eq!(self.deferred_depth,0);
        self.owned.extend(self.pending.iter().map(|request|Request {
            target:source[request.target.start..request.target.end].to_owned(),
            loader:request.loader.clone(),policy:request.policy,span:request.span,preparation:request.preparation,
        }));
        self.committed.append(&mut self.pending);
        self.undo.clear();
    }
}
