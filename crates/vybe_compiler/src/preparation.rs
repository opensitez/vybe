//! Bounded, persistent workers for preparing independent compilation units.
//! Preparation produces a common AST; linking and execution stay on the caller.

use crate::languages::Language;
use crossbeam_channel::{Receiver, Sender};
use std::collections::{HashMap, HashSet, VecDeque};
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, OnceLock};
use vybe_ast::Module;

#[cfg(test)]
#[path = "preparation_carrier_tests.rs"]
mod carrier_tests;

/// Keep generated library schemas for analysis while reusing an installed
/// definition's code. The provider validates the immutable library identity;
/// an ordinary user declaration is never converted by name alone.
pub fn bind_installed_libraries(
    module: &mut Module,
    mut installed: impl FnMut(&str, &str) -> bool,
) -> usize {
    fn bind(body: &mut [vybe_ast::Statement], namespace: &str,
        installed: &mut impl FnMut(&str, &str) -> bool,
    ) -> usize {
        let mut count = 0;
        for statement in body {
            match &mut statement.kind {
                vybe_ast::StmtKind::ClassDecl { name, modifiers, .. } => {
                    if let vybe_ast::ClassLinkage::Library { identity } = &modifiers.linkage {
                        let qualified = if namespace.is_empty() { name.clone() }
                            else { format!("{namespace}.{name}") };
                        if installed(&qualified, identity) {
                            modifiers.linkage = vybe_ast::ClassLinkage::LibraryReference {
                                identity: identity.clone(),
                            };
                            count += 1;
                        }
                    }
                }
                vybe_ast::StmtKind::NamespaceDecl { name, body } => {
                    let qualified = if namespace.is_empty() { name.clone() }
                        else { format!("{namespace}.{name}") };
                    count += bind(body, &qualified, installed);
                }
                vybe_ast::StmtKind::Block(body) => count += bind(body, namespace, installed),
                _ => {}
            }
        }
        count
    }
    bind(&mut module.body, "", &mut installed)
}

/// Retain augmentation schemas from this validated source, before importing
/// other units' schemas for lowering. Immutable output caches keep this payload
/// so later units can use the same common-AST declarations on cache replay.
pub fn source_declarations(module: &Module) -> Vec<vybe_ast::Statement> {
    fn collect(body: &[vybe_ast::Statement], out: &mut Vec<vybe_ast::Statement>) {
        for statement in body {
            match &statement.kind {
                vybe_ast::StmtKind::ClassDecl { modifiers, .. }
                    if matches!(modifiers.kind, vybe_ast::ClassKind::Trait | vybe_ast::ClassKind::Mixin | vybe_ast::ClassKind::Module) => out.push(statement.clone()),
                vybe_ast::StmtKind::NamespaceDecl { body, .. } | vybe_ast::StmtKind::Block(body) => collect(body, out),
                _ => {},
            }
        }
    }
    let mut declarations=Vec::new();collect(&module.body,&mut declarations);declarations
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LibraryDependency {
    pub name: String,
    pub identity: String,
}

/// Link requirements of this unit, independent of request-owned preparation.
pub fn required_libraries(module: &Module) -> Vec<LibraryDependency> {
    fn collect(body: &[vybe_ast::Statement], namespace: &str, out: &mut Vec<LibraryDependency>) {
        for statement in body {
            match &statement.kind {
                vybe_ast::StmtKind::ClassDecl { name, modifiers, .. } => {
                    if let vybe_ast::ClassLinkage::LibraryReference { identity } = &modifiers.linkage {
                        out.push(LibraryDependency {
                            name: if namespace.is_empty() { name.clone() } else { format!("{namespace}.{name}") },
                            identity: identity.clone(),
                        });
                    }
                }
                vybe_ast::StmtKind::NamespaceDecl { name, body } => {
                    let qualified = if namespace.is_empty() { name.clone() } else { format!("{namespace}.{name}") };
                    collect(body, &qualified, out);
                }
                vybe_ast::StmtKind::Block(body) => collect(body, namespace, out),
                _ => {}
            }
        }
    }
    let mut out = Vec::new();
    collect(&module.body, "", &mut out);
    out
}

pub struct PreparedFile {
    pub parser:vybe_runtime::registry::ParseFn,
    pub(crate) frontend_identity:crate::dynamic::FrontendIdentity,
    profile_source_value: &'static str,
    pub preparer:Option<crate::frontend_preparation::FrontendPreparer>,
    pub preparation_hints_complete:bool,
    pub statement_production:Option<crate::statement_codegen::Production>,
    pub path: PathBuf,
    pub source: String,
    pub module: Result<Module, String>,
    pub elapsed: std::time::Duration,
    pub dependencies: Option<Vec<PathBuf>>,
}

impl PreparedFile {
    pub fn matches_frontend(&self,language:Language)->bool {
        std::ptr::fn_addr_eq(self.parser,language.parse)
            && crate::dynamic::FrontendIdentity::of(language)==self.frontend_identity
            && (language.profile_source)()==self.profile_source_value
    }
    /// The caller supplies the source it opened at the actual include. A unit
    /// with other dependencies needs separate validation and stays on the
    /// normal preparation path until that validation is available.
    pub fn matches_source(&self, path: &std::path::Path, source: &str) -> bool {
        let matches = self.path == path
            && self.source == source
            && self.dependencies.as_ref().is_some_and(Vec::is_empty);
        if !matches && std::env::var_os("VYBE_PREPARATION_AUDIT").is_some() {
            eprintln!("[preparation-audit] source-rejected prepared={} actual={} path_equal={} bytes_equal={} dependencies_empty={}",
                self.path.display(), path.display(), self.path == path, self.source == source,
                self.dependencies.as_ref().is_some_and(Vec::is_empty));
        }
        matches
    }
}

/// Request-owned jobs travel with a compiled entry across compiler-service
/// handoffs. Compiled-output caches deliberately do not retain this payload.
pub struct PreparedSources(std::sync::Mutex<PreparedBatch>);

/// Frontend-owned path summaries travel with request preparation. Resolving a
/// summary may schedule source preparation, but must not evaluate user code or
/// publish declarations. Lookup supplies only currently visible string values.
pub trait DeferredPreparationHints: Send {
    fn take_ready_paths(&mut self, lookup: &mut dyn FnMut(&str) -> Option<String>) -> Vec<String>;
    fn is_empty(&self) -> bool;
}

pub(crate) type SourceHintFactory = fn(&str, &std::path::Path) -> Box<dyn DeferredPreparationHints>;

pub(crate) type DeferredHints = HashMap<PathBuf, Box<dyn DeferredPreparationHints>>;

pub(crate) type ResolvedPreparationHints = Arc<std::sync::Mutex<Vec<String>>>;

struct RecordedPreparationHints {
    inner: Box<dyn DeferredPreparationHints>,
    resolved: ResolvedPreparationHints,
}
impl DeferredPreparationHints for RecordedPreparationHints {
    fn take_ready_paths(&mut self, lookup: &mut dyn FnMut(&str) -> Option<String>) -> Vec<String> {
        let paths = self.inner.take_ready_paths(lookup);
        if let Ok(mut resolved) = self.resolved.lock() {
            // Same bound as pending speculation; this is not a parser limit.
            let remaining = 128usize.saturating_sub(resolved.len());
            resolved.extend(paths.iter().take(remaining).cloned());
        }
        paths
    }
    fn is_empty(&self) -> bool { self.inner.is_empty() }
}

#[derive(Default)]
pub(crate) struct PreparedBatch {
    pub helper_session: crate::primitives::helper_templates::Session,
    pub tickets: HashMap<PathBuf, PreparationTicket>,
    pub plan: PreparationPlan,
    pub deferred: DeferredHints,
    pub lowering_frontend: Option<Language>,
    pub lowerings: HashMap<PathBuf, crate::dynamic::LoweringTicket>,
}

/// Request-owned lookahead, independent of language syntax and VM state.
/// Bounds apply only to speculative work; declined jobs compile normally.
#[derive(Default)]
pub struct PreparationPlan {
    pending: VecDeque<(PathBuf, PathBuf, Language)>,
    queued: HashSet<PathBuf>,
    admissions: VecDeque<PathBuf>,
    // Completion-opted-in source identity facts only, never prepared payloads.
    // None is unset; a selected owner with an empty row is an explicit full clear.
    priority_owner: Option<PathBuf>,
    priority_history: VecDeque<(PathBuf, Vec<PathBuf>)>,
    // Hint admission history only. Never consulted by actual compilation.
    // None is default-off; an explicit false survives a payload handoff.
    retirement_admission: Option<bool>,
    retired: HashSet<PathBuf>,
    retired_order: VecDeque<PathBuf>,
}

impl PreparationPlan {
    /// Opt into suppressing repeated hints for completed retired work. The
    /// caller owns frontend permission. This also gates current-source ranking;
    /// disabling drops all request-local rank and retired-hint history.
    pub fn set_retirement_admission(&mut self, enabled: bool) {
        self.retirement_admission = Some(enabled);
        if !enabled {
            self.retired.clear(); self.retired_order.clear();
            self.priority_owner=None; self.priority_history.clear();
        }
    }

    fn retirement_admission_enabled(&self) -> bool {
        self.retirement_admission.unwrap_or(false)
    }

    fn is_retired(&self, key: &std::path::Path) -> bool {
        self.retirement_admission_enabled() && !self.retired.is_empty() && self.retired.contains(key)
    }

    fn remember_retired(&mut self, key: PathBuf) {
        // Same path-count bound as pending hints. Oldest eviction permits a
        // future speculative retry; it never affects an actual include.
        const CAPACITY: usize = 128;
        if !self.retirement_admission_enabled() { return; }
        if !self.retired.insert(key.clone()) { return; }
        self.retired_order.push_back(key.clone());
        self.pending.retain(|(pending, _, _)| pending != &key);
        self.queued.remove(&key);
        if self.retired_order.len() > CAPACITY {
            self.retired.remove(&self.retired_order.pop_front().unwrap());
        }
    }

    pub fn len(&self) -> usize { self.pending.len() }
    pub fn is_empty(&self) -> bool { self.pending.is_empty() }
    pub(crate) fn extend(&mut self, other: Self) {
        self.admissions.extend(other.admissions);
        if self.retirement_admission.is_none() {
            self.retirement_admission = other.retirement_admission;
        }
        if self.retirement_admission_enabled() {
            if self.priority_owner.is_none() { self.priority_owner=other.priority_owner; }
            for row in other.priority_history {
                // Receiving full-empty and incremental rows remain authoritative.
                if !self.priority_history.iter().any(|(owner,_)|owner==&row.0) {
                    self.priority_history.push_back(row);
                }
            }
            self.trim_priority_history();
        }
        for key in other.retired_order { self.remember_retired(key); }
        for (key, path, language) in other.pending { self.enqueue_as(key, path, language); }
    }

    pub fn enqueue(&mut self, path: PathBuf, language: Language) -> bool {
        self.enqueue_as(path.clone(), path, language)
    }

    // File identity deduplicates work; parser-visible spelling preserves magic
    // file constants and source diagnostics.
    pub(crate) fn enqueue_as(&mut self, key: PathBuf, path: PathBuf, language: Language) -> bool {
        if self.is_retired(&key) || self.pending.len() >= 128 || !self.queued.insert(key.clone()) { return false; }
        self.pending.push_back((key, path, language));
        true
    }

    /// Prioritize paths discovered in the currently executing source over
    /// speculative descendants. Preserve source order and the pending bound;
    /// displaced speculation still has ordinary foreground compilation.
    pub(crate) fn prioritize_as(&mut self,
        paths: impl IntoIterator<Item = (PathBuf, PathBuf, Language)>,
    ) {
        let mut seen = HashSet::new();
        let priority: Vec<_> = paths.into_iter()
            .filter(|(key, _, _)| !self.is_retired(key) && seen.insert(key.clone())).take(128).collect();
        self.pending.retain(|(key, _, _)| !seen.contains(key));
        self.queued.retain(|key| !seen.contains(key));
        while self.pending.len() + priority.len() > 128 {
            if let Some((key, _, _)) = self.pending.pop_back() { self.queued.remove(&key); }
        }
        for (key, path, language) in priority.into_iter().rev() {
            self.queued.insert(key.clone());
            self.pending.push_front((key, path, language));
        }
    }

    /// Full receipts replace this owner's known siblings, including an explicit
    /// empty receipt. Incremental refresh merges newly resolved identities; an
    /// empty refresh preserves retained siblings and restores a returning owner.
    pub(crate) fn prioritize_current(&mut self, owner:PathBuf, paths:&[PathBuf], full:bool) {
        if !self.retirement_admission_enabled() { return; }
        let mut row=self.priority_history.iter().position(|(key,_)|key==&owner)
            .and_then(|index|self.priority_history.remove(index))
            .unwrap_or_else(||(owner.clone(),Vec::new()));
        if full { row.1.clear(); }
        for path in paths {
            if row.1.len()==128 { break; }
            if !row.1.contains(path) { row.1.push(path.clone()); }
        }
        self.priority_owner=Some(owner);
        self.priority_history.push_back(row);
        self.trim_priority_history();
    }

    fn trim_priority_history(&mut self) {
        while self.priority_history.len()>128
            ||self.priority_history.iter().map(|(_,paths)|paths.len()).sum::<usize>()>128 {
            // Keep the selected owner's row. Older owner facts are expendable;
            // a returning evicted owner falls back to FIFO until rediscovered.
            let Some(index)=self.priority_history.iter().position(|(owner,_)|
                self.priority_owner.as_ref()!=Some(owner)) else { break; };
            self.priority_history.remove(index);
        }
    }

    fn current_priority(&self)->&[PathBuf] {
        if !self.retirement_admission_enabled() { return &[]; }
        self.priority_owner.as_ref().and_then(|owner|
            self.priority_history.iter().find(|(key,_)|key==owner))
            .map_or(&[],|(_,paths)|paths.as_slice())
    }

    // Gate-only diagnostic of the real private admission predicate.
    #[cfg(any(test,v6_service_gate))]
    pub(crate) fn retirement_rank_gate_enabled(&self)->bool {
        self.retirement_admission_enabled()
    }

    /// Retire at most one completed speculative unit when later work is pending.
    /// The caller must drop its paired lowering receiver before refilling.
    /// Running jobs and the actual requested unit retain ownership. No requeue.
    pub(crate) fn retire_completed_for_pending(
        &mut self, active: &mut HashMap<PathBuf, PreparationTicket>,
        protected: Option<&std::path::Path>,
    ) -> Option<(PathBuf, PreparationTicket)> {
        self.admissions.retain(|key| active.contains_key(key));
        if self.pending.is_empty() { return None; }
        let ready=|key:&PathBuf|protected!=Some(key.as_path())
            &&active.get(key).is_some_and(PreparationTicket::is_ready);
        let priority=self.current_priority();
        let index=if priority.is_empty() { self.admissions.iter().position(ready) }
            else { self.admissions.iter().position(|key|ready(key)&&!priority.contains(key))
                .or_else(||self.admissions.iter().enumerate().filter(|(_,key)|ready(key))
                    .filter_map(|(index,key)|priority.iter().position(|p|p==key).map(|rank|(index,rank)))
                    .max_by_key(|(_,rank)|*rank).map(|(index,_)|index)) }?;
        let key = self.admissions.remove(index).unwrap();
        let ticket = active.remove(&key)?;
        self.remember_retired(key.clone());
        Some((key, ticket))
    }

    /// Fill a small active window whenever the caller consumes or cancels a
    /// ticket. A full shared pool keeps the candidate for a later attempt.
    pub fn replenish(&mut self, pool: &PreparationPool,
        active: &mut HashMap<PathBuf, PreparationTicket>, window: usize,
        skip: impl FnMut(&std::path::Path) -> bool,
    ) {
        self.replenish_with_hints(pool, active, window, None, skip);
    }

    pub(crate) fn replenish_with_hints(&mut self, pool: &PreparationPool,
        active: &mut HashMap<PathBuf, PreparationTicket>, window: usize,
        hints: Option<SourceHintFactory>, mut skip: impl FnMut(&std::path::Path) -> bool,
    ) {
        self.replenish_by(active,window,|_key,path,language|pool.submit_with_hints(path,language,hints),skip);
    }
    pub(crate) fn replenish_by(&mut self,active:&mut HashMap<PathBuf,PreparationTicket>,window:usize,
        mut submit:impl FnMut(PathBuf,PathBuf,Language)->Option<PreparationTicket>,mut skip:impl FnMut(&std::path::Path)->bool,
    ){
        self.admissions.retain(|key| active.contains_key(key));
        while active.len() < window {
            let Some((key, path, language)) = self.pending.pop_front() else { break; };
            if active.contains_key(&key) || skip(&key) {
                self.queued.remove(&key);
                continue;
            }
            if let Some(ticket) = submit(key.clone(), path.clone(), language) {
                self.queued.remove(&key);
                self.admissions.push_back(key.clone());
                active.insert(key, ticket);
            } else {
                self.pending.push_front((key, path, language));
                break;
            }
        }
    }
}

impl std::fmt::Debug for PreparedSources {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let pending = self.0.lock().unwrap();
        f.debug_struct("PreparedSources")
            .field("pending", &pending.tickets.len())
            .field("planned", &pending.plan.len())
            .field("deferred", &pending.deferred.len())
            .finish()
    }
}

impl PreparedSources {
    pub(crate) fn new(tickets: HashMap<PathBuf, PreparationTicket>, plan: PreparationPlan,
        deferred: DeferredHints,
    ) -> Self {
        Self(std::sync::Mutex::new(PreparedBatch { tickets, plan, deferred, ..Default::default() }))
    }
    /// Keep worker-result ownership alive when entry compilation and execution
    /// use different compiler services. The payload remains one-shot.
    pub(crate) fn with_lowerings(self, frontend: Language,
        lowerings: HashMap<PathBuf, crate::dynamic::LoweringTicket>,
    ) -> Self {
        {
            let mut batch = self.0.lock().unwrap();
            batch.lowering_frontend = Some(frontend);
            batch.lowerings = lowerings;
        }
        self
    }
    pub(crate) fn with_helper_session(self, session: crate::primitives::helper_templates::Session) -> Self {
        self.0.lock().unwrap().helper_session = session;
        self
    }
    pub(crate) fn take(&self) -> PreparedBatch {
        std::mem::take(&mut *self.0.lock().unwrap())
    }
}

pub(crate) type PostPrepare = Box<dyn FnOnce(&mut PreparedFile) + Send>;

struct Job {
    path: PathBuf,
    language: Language,
    preparer:Option<crate::frontend_preparation::FrontendPreparer>,
    frontend_identity:crate::dynamic::FrontendIdentity,
    profile_source_value: &'static str,
    cancelled: Arc<AtomicBool>,
    result: Sender<Option<PreparedFile>>,
    hints: Option<(Sender<(PathBuf, Box<dyn DeferredPreparationHints>)>, Option<SourceHintFactory>,
        Option<ResolvedPreparationHints>)>,
    post_prepare: Option<PostPrepare>,
}

pub struct PreparationTicket {
    result: Receiver<Option<PreparedFile>>,
    cancelled: Arc<AtomicBool>,
    hints: Option<Receiver<(PathBuf, Box<dyn DeferredPreparationHints>)>>,
    resolved_hints: Option<ResolvedPreparationHints>,
}

impl PreparationTicket {
    pub(crate) fn resolved_hints(&self) -> Option<ResolvedPreparationHints> {
        self.resolved_hints.clone()
    }
    pub(crate) fn has_source_hints(&self) -> bool { self.hints.is_some() }
    pub(crate) fn hint_receiver(&self)->Option<&Receiver<(PathBuf,Box<dyn DeferredPreparationHints>)>> {self.hints.as_ref()}
    /// Pure discovery is available before this unit's parser finishes. The
    /// request's compiler decides whether to schedule any discovered paths;
    /// workers never publish declarations or read request bindings.
    pub(crate) fn take_hints(&self) -> Option<(PathBuf, Box<dyn DeferredPreparationHints>)> {
        self.hints.as_ref()?.try_recv().ok()
    }
    pub fn is_ready(&self) -> bool {
        !self.result.is_empty()
    }

    pub fn wait(self) -> Option<PreparedFile> {
        self.result.recv().ok().flatten()
    }

    #[cfg(test)]
    fn wait_with_hints(self) -> (Option<PreparedFile>, Option<(PathBuf, Box<dyn DeferredPreparationHints>)>) {
        let prepared = self.result.recv().ok().flatten();
        (prepared, self.take_hints())
    }

    pub(crate) fn wait_with_early_hints(self,
        mut discover: impl FnMut((PathBuf, Box<dyn DeferredPreparationHints>)),
    ) -> Option<PreparedFile> {
        if let Some(hints)=&self.hints {
            loop {
                crossbeam_channel::select! {
                    recv(hints)->hint=>match hint {
                        Ok(hint)=>discover(hint),
                        Err(_)=>{let prepared=self.result.recv().ok().flatten();while let Some(hint)=self.take_hints(){discover(hint);}return prepared;}
                    },
                    recv(self.result)->prepared=>{
                        while let Some(hint)=self.take_hints(){discover(hint);}
                        return prepared.ok().flatten();
                    }
                }
            }
        }
        self.result.recv().ok().flatten()
    }
}

impl Drop for PreparationTicket {
    fn drop(&mut self) {
        self.cancelled.store(true, Ordering::Relaxed);
    }
}

enum Work { Prepare(Job), Lower(Box<dyn FnOnce() + Send>) }

pub struct PreparationPool {
    jobs: Sender<Work>,
}

impl PreparationPool {
    pub fn new(workers: usize, queue_capacity: usize) -> std::io::Result<Self> {
        if workers == 0 || queue_capacity == 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "preparation workers and queue capacity must be positive",
            ));
        }
        let (jobs, receiver) = crossbeam_channel::bounded::<Work>(queue_capacity);
        for index in 0..workers {
            let receiver = receiver.clone();
            std::thread::Builder::new()
                .name(format!("vybe-prepare-{index}"))
                .stack_size(64 * 1024 * 1024)
                .spawn(move || {
                    while let Ok(work) = receiver.recv() {
                        let job=match work {
                            Work::Prepare(job)=>job,
                            Work::Lower(lower)=>{let _=std::panic::catch_unwind(std::panic::AssertUnwindSafe(lower));continue;}
                        };
                        let _job_span = crate::compilation_timeline::span("prepare_job", &job.path);
                        if job.cancelled.load(Ordering::Relaxed) {
                            continue;
                        }
                        let prepared =
                            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                                let started = std::time::Instant::now();
                                let source = std::fs::read_to_string(&job.path).ok()?;
                                if job.cancelled.load(Ordering::Relaxed) {
                                    return None;
                                }
                                if let Some((sender, Some(discover), resolved)) = &job.hints {
                                    let mut hints = discover(&source, &job.path);
                                    if let Some(resolved) = resolved {
                                        hints = Box::new(RecordedPreparationHints { inner: hints, resolved: resolved.clone() });
                                    }
                                    let sent = sender.try_send((job.path.clone(), hints)).is_ok();
                                    crate::compilation_timeline::event("source_hints", if sent { "sent" } else { "declined" }, &job.path);
                                }
                                let mut bundle = crate::dynamic::bundle_from_source(
                                    source,
                                    job.language,
                                    job.path.clone(),
                                );
                                let parse_span = crate::compilation_timeline::span("worker_frontend", &job.path);
                                let (result, dependencies) = crate::bundle::record_source_reads(||
                                    bundle.prepared_frontend_with_sink(None,job.preparer,&mut |mut hints| {
                                        if !job.cancelled.load(Ordering::Relaxed) {
                                            if let Some((sender,_,resolved))=&job.hints {
                                                if let Some(resolved)=resolved {hints=Box::new(RecordedPreparationHints{inner:hints,resolved:resolved.clone()});}
                                                let sent=sender.try_send((job.path.clone(),hints)).is_ok();
                                                crate::compilation_timeline::event("parser_hints", if sent { "sent" } else { "declined" }, &job.path);
                                            }
                                        }
                                    }));
                                drop(parse_span);
                                let mut complete=false;
                                let mut statement_production=None;
                                let module=match result {
                                    Ok(mut result)=>{
                                        complete=result.preparation_hints_complete;
                                        statement_production=result.statement_production.take();
                                        if !job.cancelled.load(Ordering::Relaxed) {
                                            if let (Some((sender,_,resolved)),Some(mut hints))=(&job.hints,result.hints.take()) {
                                                if let Some(resolved)=resolved {hints=Box::new(RecordedPreparationHints{inner:hints,resolved:resolved.clone()});}
                                                // At most one early scout and one committed parse receipt.
                                                // Registered jobs have a bounded two-message mailbox.
                                                let _=sender.try_send((job.path.clone(),hints));
                                            }
                                        }
                                        Ok(result.module)
                                    },Err(error)=>Err(error)
                                };
                                let mut prepared=PreparedFile {
                                    parser:job.language.parse,preparer:job.preparer,frontend_identity:job.frontend_identity,profile_source_value:job.profile_source_value,preparation_hints_complete:complete,statement_production,
                                    path: job.path,
                                    source: std::mem::take(&mut bundle.sources[0].code),
                                    module,
                                    elapsed: started.elapsed(),
                                    dependencies,
                                };
                                if !job.cancelled.load(Ordering::Relaxed) {
                                    if let Some(post)=job.post_prepare {
                                        let _lowering_span = crate::compilation_timeline::span("worker_post_prepare", &prepared.path);
                                        // A lowering panic must not discard an already parsed source.
                                        let _=std::panic::catch_unwind(std::panic::AssertUnwindSafe(||post(&mut prepared)));
                                    }
                                }
                                Some(prepared)
                            }))
                            .ok()
                            .flatten();
                        let _ = job.result.send(prepared);
                    }
                })?;
        }
        Ok(Self { jobs })
    }

    /// Queue without blocking the executing VM. A full queue simply declines
    /// speculation; the normal synchronous preparation remains available.
    pub fn submit(&self, path: PathBuf, language: Language) -> Option<PreparationTicket> {
        self.submit_with_hints(path, language, None)
    }

    pub(crate) fn submit_with_hints(&self, path: PathBuf, language: Language,
        discover: Option<SourceHintFactory>,
    ) -> Option<PreparationTicket> {
        self.submit_with_post_prepare(path,language,discover,None)
    }

    pub(crate) fn submit_with_post_prepare(&self,path:PathBuf,language:Language,
        discover:Option<SourceHintFactory>,post_prepare:Option<PostPrepare>,
    )->Option<PreparationTicket>{
        let (sender, result) = crossbeam_channel::bounded(1);
        let registration=crate::frontend_preparation::snapshot(language.parse);
        let preparer=registration.preparer();
        let discover=if preparer.is_some_and(|hook|!hook.early_hints){None}else{discover};
        let (hints_sender, hints, resolved_hints) = if discover.is_some()||preparer.is_some() {
            let (sender, receiver) = crossbeam_channel::bounded(if preparer.is_some(){2}else{1});
            let resolved = post_prepare.is_some().then(|| Arc::new(std::sync::Mutex::new(Vec::new())));
            (Some((sender, discover, resolved.clone())), Some(receiver), resolved)
        } else { (None, None, None) };
        let cancelled = Arc::new(AtomicBool::new(false));
        crate::compilation_timeline::event("prepare_queue", "submit", &path);
        let queued = self.jobs
            .try_send(Work::Prepare(Job {
                path,
                language,preparer,frontend_identity:crate::dynamic::FrontendIdentity::with_registration(language,registration),profile_source_value:(language.profile_source)(),
                cancelled: cancelled.clone(),
                result: sender,
                hints: hints_sender,
                post_prepare,
            }));
        if let Err(error) = queued {
            if let Work::Prepare(job) = error.into_inner() {
                crate::compilation_timeline::event("prepare_queue", "declined", &job.path);
            }
            return None;
        }
        Some(PreparationTicket { result, cancelled, hints, resolved_hints })
    }

    pub(crate) fn submit_lowering(&self, work: Box<dyn FnOnce()+Send>) -> bool {
        self.jobs.try_send(Work::Lower(work)).is_ok()
    }

    /// Workers share no VM or compilation results. Each ticket owns its data,
    /// so keeping the threads warm does not retain tenant state between runs.
    pub(crate) fn shared() -> Option<&'static Self> {
        static POOL: OnceLock<Option<PreparationPool>> = OnceLock::new();
        POOL.get_or_init(|| Self::new(2, 4).ok()).as_ref()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::{Condvar, Mutex};

    fn admission_test_language() -> Language {
        Language { name: "admission-test", parse: |_| Err("unused parser".into()),
            profile_source: || "", emit_dispatch: None, normalize_class: None,
            register_tree: None, expand_source: None }
    }

    #[test]
    fn retired_hint_suppression_requires_explicit_opt_in() {
        let key = PathBuf::from("retired");
        let mut legacy = PreparationPlan::default();
        legacy.remember_retired(key.clone());
        assert!(legacy.enqueue(key.clone(), admission_test_language()));
        let mut optimized = PreparationPlan::default();
        optimized.set_retirement_admission(true);
        optimized.remember_retired(key.clone());
        assert!(!optimized.enqueue(key.clone(), admission_test_language()));
        optimized.set_retirement_admission(false);
        assert!(optimized.enqueue(key, admission_test_language()));
    }

    #[test]
    fn explicit_admission_opt_out_survives_source_handoff() {
        let key = PathBuf::from("retired");
        let mut source = PreparationPlan::default();
        source.set_retirement_admission(true);
        source.remember_retired(key.clone());
        let mut destination = PreparationPlan::default();
        destination.set_retirement_admission(false);
        destination.extend(source);
        assert!(destination.enqueue(key, admission_test_language()));
        assert!(destination.retired.is_empty());
    }

    #[test]
    fn retired_hint_history_is_bounded_and_follows_enabled_handoff() {
        let mut source = PreparationPlan::default();
        source.set_retirement_admission(true);
        for index in 0..129 { source.remember_retired(PathBuf::from(format!("unit{index}"))); }
        let mut destination = PreparationPlan::default();
        destination.extend(source);
        assert_eq!(destination.retired.len(), 128);
        assert!(destination.enqueue(PathBuf::from("unit0"), admission_test_language()));
        assert!(!destination.enqueue(PathBuf::from("unit128"), admission_test_language()));
    }

    #[test]
    fn resolved_hint_history_survives_consumption_without_limiting_discovery() {
        struct Hints(Vec<String>);
        impl DeferredPreparationHints for Hints {
            fn take_ready_paths(&mut self, _: &mut dyn FnMut(&str) -> Option<String>) -> Vec<String> {
                std::mem::take(&mut self.0)
            }
            fn is_empty(&self) -> bool { self.0.is_empty() }
        }
        let resolved = Arc::new(Mutex::new(Vec::new()));
        let paths: Vec<_> = (0..130).map(|i| format!("child{i}")).collect();
        let mut hints = RecordedPreparationHints { inner: Box::new(Hints(paths.clone())), resolved: resolved.clone() };
        assert_eq!(hints.take_ready_paths(&mut |_| None), paths);
        assert!(hints.is_empty());
        drop(hints);
        assert_eq!(*resolved.lock().unwrap(), paths[..128]);
    }

    #[test]
    fn awaited_preparation_preserves_discovery_that_arrives_after_polling() {
        fn parse(source: &str) -> Result<Module, String> {
            Ok(Module { name: source.into(), language: vybe_ast::Lang::Unknown,
                body: vec![], imports: vec![], directives: Default::default(), canon: Default::default() })
        }
        let language = Language { name: "hint-test", parse, profile_source: || "",
            emit_dispatch: None, normalize_class: None, register_tree: None, expand_source: None };
        let path = std::env::temp_dir().join(format!("vybe-late-hints-{}", uuid::Uuid::new_v4()));
        std::fs::write(&path, "require 'child.php';").unwrap();
        let pool = PreparationPool::new(1, 1).unwrap();
        let ticket = pool.submit_with_hints(path.clone(), language, Some(crate::php_include_hints::discover)).unwrap();
        let (prepared, hints) = ticket.wait_with_hints();
        assert!(prepared.unwrap().module.is_ok());
        let (caller, mut hints) = hints.expect("unconsumed discovery survives awaiting AST");
        assert_eq!(caller, path);
        assert_eq!(hints.take_ready_paths(&mut |_| None), ["child.php"]);
        std::fs::remove_file(path).unwrap();
    }

    #[test]
    fn source_hints_arrive_while_parent_parse_is_still_blocked() {
        static PARSE_GATE: OnceLock<(Mutex<(bool, bool)>, Condvar)> = OnceLock::new();
        fn parse(source: &str) -> Result<Module, String> {
            let (lock, signal) = PARSE_GATE.get_or_init(|| (Mutex::new((false, false)), Condvar::new()));
            let mut state = lock.lock().unwrap();
            if source.contains("PARENT") {
                state.0 = true;
                signal.notify_all();
                let (next, timeout) = signal.wait_timeout_while(state,
                    std::time::Duration::from_secs(10), |state| !state.1).unwrap();
                assert!(!timeout.timed_out(), "grandchild must parse before parent completes");
                state = next;
            } else {
                state.1 = true;
                signal.notify_all();
            }
            drop(state);
            Ok(Module { name: source.into(), language: vybe_ast::Lang::Unknown,
                body: vec![], imports: vec![], directives: Default::default(), canon: Default::default() })
        }
        let language = Language { name: "php", parse, profile_source: || "",
            emit_dispatch: None, normalize_class: None, register_tree: None, expand_source: None };
        let base = std::env::temp_dir().join(format!("vybe-early-hints-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let parent = base.join("parent.php");
        let child = base.join("child.php");
        std::fs::write(&parent, "<?php /* PARENT */ require __DIR__ . '/child.php';").unwrap();
        std::fs::write(&child, "<?php return 7;").unwrap();
        let pool = PreparationPool::new(2, 4).unwrap();
        let ticket = pool.submit_with_hints(parent.clone(), language,
            Some(crate::php_include_hints::discover)).unwrap();
        let (lock, signal) = PARSE_GATE.get_or_init(|| (Mutex::new((false, false)), Condvar::new()));
        let (state, timeout) = signal.wait_timeout_while(lock.lock().unwrap(),
            std::time::Duration::from_secs(10), |state| !state.0).unwrap();
        assert!(!timeout.timed_out());
        drop(state);
        assert!(!ticket.is_ready(), "source discovery precedes completed parent AST");
        let mut discovered = 0;
        let prepared_parent = ticket.wait_with_early_hints(|(caller, mut hints)| {
            discovered += 1;
            assert_eq!(caller, parent);
            let paths = hints.take_ready_paths(&mut |_| None);
            assert_eq!(paths, [child.to_string_lossy().into_owned()]);
            let prepared_child = pool.submit(child.clone(), language).unwrap().wait().unwrap();
            assert!(prepared_child.module.is_ok());
        }).unwrap();
        assert_eq!(discovered, 1, "request consumes discovery once");
        assert!(prepared_parent.module.is_ok());
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn deferred_hints_survive_handoff_and_resolve_current_request_values_once() {
        struct Hint(bool);
        impl DeferredPreparationHints for Hint {
            fn take_ready_paths(&mut self, lookup: &mut dyn FnMut(&str) -> Option<String>) -> Vec<String> {
                if !self.0 { return vec![]; }
                let Some(root) = lookup("ROOT") else { return vec![]; };
                self.0 = false;
                vec![format!("{root}/child")]
            }
            fn is_empty(&self) -> bool { !self.0 }
        }
        let path = PathBuf::from("/source/parent");
        let mut deferred: DeferredHints = HashMap::new();
        deferred.insert(path.clone(), Box::new(Hint(true)));
        let shared = Arc::new(PreparedSources::new(Default::default(), Default::default(), deferred));
        let mut batch = shared.take();
        assert!(shared.take().deferred.is_empty(), "handoff has one consumer");
        let hints = batch.deferred.get_mut(&path).unwrap();
        assert!(hints.take_ready_paths(&mut |_| None).is_empty());
        assert!(!hints.is_empty());
        assert_eq!(hints.take_ready_paths(&mut |_| Some("/request-a".into())), ["/request-a/child"]);
        assert!(hints.take_ready_paths(&mut |_| Some("/request-b".into())).is_empty());
        assert!(hints.is_empty());
        assert_eq!(Hint(true).take_ready_paths(&mut |_| Some("/request-b".into())), ["/request-b/child"]);
    }

    #[test]
    fn library_binding_requires_identity_and_preserves_user_declarations_and_schemas() {
        use vybe_ast::{ClassLinkage, ClassModifiers, Statement, StmtKind};
        fn declaration(linkage: ClassLinkage) -> Statement {
            Statement::new(StmtKind::ClassDecl {
                name: "Shared".into(), parents: vec!["Base".into()], interfaces: vec![],
                members: vec![vybe_ast::ClassMember::Field {
                    name: "state".into(), type_hint: Some("int".into()),
                    init: Some(vybe_ast::Expression::int(7)),
                    modifiers: Default::default(), with_events: false,
                    array_bounds: None, storage: None,
                }],
                modifiers: ClassModifiers { linkage, ..Default::default() }, decorators: vec![],
            })
        }
        let mut module = Module { name: "generic-library".into(), language: vybe_ast::Lang::Unknown,
            body: vec![declaration(ClassLinkage::Definition),
                declaration(ClassLinkage::Library { identity: "old-version".into() }),
                Statement::new(StmtKind::NamespaceDecl { name: "Tools".into(),
                    body: vec![declaration(ClassLinkage::Library { identity: "installed-version".into() })] })],
            imports: vec![], directives: Default::default(), canon: Default::default() };
        let rebound = bind_installed_libraries(&mut module, |name, identity| {
            name == "Tools.Shared" && identity == "installed-version"
        });
        assert_eq!(rebound, 1);
        let StmtKind::ClassDecl { modifiers, .. } = &module.body[0].kind else { panic!() };
        assert!(matches!(modifiers.linkage, ClassLinkage::Definition));
        let StmtKind::ClassDecl { modifiers, .. } = &module.body[1].kind else { panic!() };
        assert!(matches!(modifiers.linkage, ClassLinkage::Library { .. }));
        let StmtKind::NamespaceDecl { body, .. } = &module.body[2].kind else { panic!() };
        let StmtKind::ClassDecl { members, parents, modifiers, .. } = &body[0].kind else { panic!() };
        assert!(matches!(modifiers.linkage, ClassLinkage::LibraryReference { .. }));
        assert_eq!(parents, &["Base"]);
        assert!(matches!(&members[0], vybe_ast::ClassMember::Field { init: Some(_), .. }));
    }

    #[test]
    fn current_source_priority_preserves_order_deduplication_and_pending_bound() {
        fn parse(_: &str) -> Result<Module, String> { unreachable!() }
        let language = Language { name: "priority-test", parse, profile_source: || "",
            emit_dispatch: None, normalize_class: None, register_tree: None, expand_source: None };
        let mut plan = PreparationPlan::default();
        for i in 0..128 { assert!(plan.enqueue(PathBuf::from(format!("cold{i}")), language)); }
        plan.prioritize_as([
            (PathBuf::from("cold100"), PathBuf::from("spelling100"), language),
            (PathBuf::from("next"), PathBuf::from("next"), language),
            (PathBuf::from("cold100"), PathBuf::from("duplicate"), language),
        ]);
        assert_eq!(plan.len(), 128);
        assert_eq!(plan.pending[0].0, PathBuf::from("cold100"));
        assert_eq!(plan.pending[0].1, PathBuf::from("spelling100"));
        assert_eq!(plan.pending[1].0, PathBuf::from("next"));
        assert_eq!(plan.pending[2].0, PathBuf::from("cold0"));
        assert!(!plan.queued.contains(&PathBuf::from("cold127")));
        assert_eq!(plan.pending.len(), plan.queued.len());
        let mut transferred = PreparationPlan::default();
        transferred.extend(plan);
        assert_eq!(transferred.pending[0].0, PathBuf::from("cold100"));
        assert_eq!(transferred.pending[1].0, PathBuf::from("next"));
    }

    #[test]
    fn preparation_plan_refills_and_survives_service_handoff_without_duplicates() {
        fn parse(source: &str) -> Result<Module, String> {
            Ok(Module { name: source.into(), language: vybe_ast::Lang::Unknown,
                body: vec![], imports: vec![], directives: Default::default(), canon: Default::default() })
        }
        let language = Language { name: "plan-test", parse, profile_source: || "",
            emit_dispatch: None, normalize_class: None, register_tree: None, expand_source: None };
        let base = std::env::temp_dir().join(format!("vybe-plan-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let paths: Vec<_> = (0..8).map(|i| {
            let path = base.join(format!("unit_{i}"));
            std::fs::write(&path, format!("unit_{i}")).unwrap();
            path
        }).collect();
        let pool = PreparationPool::new(2, 8).unwrap();
        let mut plan = PreparationPlan::default();
        for path in &paths {
            assert!(plan.enqueue(path.clone(), language));
            assert!(!plan.enqueue(path.clone(), language));
        }
        let mut active = HashMap::new();
        plan.replenish(&pool, &mut active, 4, |_| false);
        assert_eq!(active.len(), 4);
        assert_eq!(plan.len(), 4);
        let handoff = PreparedSources::new(active, plan, Default::default());
        let PreparedBatch { mut tickets, mut plan, .. } = handoff.take();
        assert!(handoff.take().tickets.is_empty(), "request jobs have one consumer");
        for (i, path) in paths.iter().enumerate() {
            let ticket = tickets.remove(path).expect("next unit already scheduled");
            plan.replenish(&pool, &mut tickets, 4, |_| false);
            let prepared = ticket.wait().expect("prepared result");
            assert_eq!(prepared.module.unwrap().name, format!("unit_{i}"));
            assert!(tickets.len() <= 4);
        }
        assert!(plan.is_empty() && tickets.is_empty());
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn independent_units_prepare_concurrently_and_keep_their_results_separate() {
        fn parse(source: &str) -> Result<Module, String> {
            static ARRIVALS: OnceLock<(Mutex<usize>, Condvar)> = OnceLock::new();
            let (lock, ready) = ARRIVALS.get_or_init(|| (Mutex::new(0), Condvar::new()));
            let mut arrived = lock.lock().unwrap();
            *arrived += 1;
            ready.notify_all();
            let (arrived, timeout) = ready
                .wait_timeout_while(arrived, std::time::Duration::from_secs(10), |count| {
                    *count < 2
                })
                .unwrap();
            assert!(
                !timeout.timed_out() && *arrived == 2,
                "both parses must run concurrently"
            );
            Ok(Module {
                name: source.to_string(),
                language: vybe_ast::Lang::Unknown,
                body: vec![],
                imports: vec![],
                directives: Default::default(),
                canon: Default::default(),
            })
        }
        let language = Language {
            name: "preparation-test",
            parse,
            profile_source: || "",
            emit_dispatch: None,
            normalize_class: None,
            register_tree: None,
            expand_source: None,
        };
        let base = std::env::temp_dir().join(format!("vybe-preparation-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let first = base.join("first");
        let second = base.join("second");
        std::fs::write(&first, "first unit").unwrap();
        std::fs::write(&second, "second unit").unwrap();
        let pool = PreparationPool::new(2, 4).unwrap();
        let a = pool.submit(first.clone(), language).unwrap();
        let b = pool.submit(second, language).unwrap();
        let a = a.wait().unwrap();
        let b = b.wait().unwrap();
        assert!(a.matches_source(&first, "first unit"));
        assert!(!a.matches_source(&first, "edited unit"));
        assert!(!a.matches_source(&base.join("other"), "first unit"));
        assert_eq!(a.module.unwrap().name, "first unit");
        assert_eq!(b.module.unwrap().name, "second unit");
        assert_eq!(a.source, "first unit");
        assert_eq!(b.source, "second unit");
        std::fs::remove_dir_all(base).unwrap();
    }
    include!("preparation_retirement_priority_tests.rs");
}
