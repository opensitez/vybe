mod request_code_space;
use request_code_space::{ResidentEntry,SourceReceipt};
use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};

use crate::bundle::{Bundle, CompiledBundle, EntryPoint, SourceFile};
use crate::languages::{self, Language};
use crate::primitives::HostImportMetadata;
use vybe_ast::{ClassKind, ClassMember, Module, Statement, StmtKind};
use vybe_runtime::capabilities::{Capabilities, Capability};
use vybe_runtime::chunk::Chunk;
use vybe_runtime::chunk::Import;
use vybe_runtime::debugger::{DebugPhase, DebugReportScope};
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{HostContext, ImportTarget, VM, Value};

thread_local! {
    static ACTIVE_PHP_RUNTIME: RefCell<Option<*mut PhpIncludeRuntime>> = const { RefCell::new(None) };
    static ACTIVE_JS_RUNTIME: RefCell<Option<*mut JsDynamicRuntime>> = const { RefCell::new(None) };
    static CONSTRUCTED_JS_FUNCTIONS: RefCell<HashMap<u64, *mut ConstructedJsFunction>> = RefCell::new(HashMap::new());
}

static NEXT_JS_DYNAMIC_ID: AtomicU64 = AtomicU64::new(1);

/// PHP's language adapter can also run without an active dynamic compiler.
/// In that case module variables use the ordinary request-global store.
pub fn php_include_scope() -> Value {
    ACTIVE_PHP_RUNTIME.with(|slot| {
        slot.borrow().as_ref().and_then(|pointer| unsafe { (**pointer).include_scope.clone() })
            .unwrap_or(Value::Null)
    })
}

/// Drop every `new Function(...)` backing state and restart the id counter.
///
/// ⛔ THIS FILE WAS MISSED BY THE STATIC-STATE CONVERSION. Every other walker
/// owns its per-compilation registries in one struct built by `parse` and
/// dropped when it returns (`PyWalker` — "71 process-global statics across
/// eight `thread_local!` blocks"; the php and wast walkers say the same). These
/// two could not follow that pattern, because a dynamic function's state must
/// OUTLIVE the compilation — the host fn registered below calls into it at run
/// time — so its real owner is the VM, not the compile.
///
/// Nothing reclaimed it. `CONSTRUCTED_JS_FUNCTIONS` grew a `Box::into_raw`
/// entry per `new Function(...)`, each holding a WHOLE `VM`, and no site ever
/// removed one: a warm worker leaked every dynamic function every program had
/// ever built, for the life of the process. `NEXT_JS_DYNAMIC_ID` never
/// restarted either, so `__vybe_dynamic_function_<n>` depended on how many
/// dynamic functions had been compiled BEFORE — the same source lowered to
/// different names as the 1st and the 100th program on a worker. That is
/// exactly the defect the python conversion called out in its own comment.
///
/// ⚠ SAFETY. The first version of this comment said the closures are dropped
/// by `VM::reset_to` before this runs. THAT IS FALSE — `host_fns` is a `Vec`
/// that `reset_to` never truncates, so every host closure a program registers
/// outlives the reset. The real argument is narrower and does hold: a pointer
/// is reachable ONLY through this map, `drain` removes and yields it in one
/// step, and the surviving closure re-reads the map on every call. After a
/// reset it finds no entry and takes the `backing state missing for id` path
/// that already exists — an error, not a dangling deref. Nothing else holds a
/// copy, so freeing the drained box is sound.
///
/// ⛔ Do not "simplify" this by keeping the pointer anywhere outside the map.
/// The single point of reachability is the whole proof.
pub fn reset_dynamic_compilation_state() {
    CONSTRUCTED_JS_FUNCTIONS.with(|slot| {
        for (_, ptr) in slot.borrow_mut().drain() {
            if !ptr.is_null() {
                // Reclaims the `Box::into_raw` at the registration site.
                drop(unsafe { Box::from_raw(ptr) });
            }
        }
    });
    NEXT_JS_DYNAMIC_ID.store(1, Ordering::Relaxed);
}

#[derive(Clone,Copy,Debug,PartialEq,Eq,Hash)]
pub(crate) struct FrontendIdentity([usize;26]);
impl FrontendIdentity {
    pub(crate) fn of(language:Language)->Self {Self::with_registration(language,crate::frontend_preparation::snapshot(language.parse))}
    pub(crate) fn with_registration(language:Language,registration:crate::frontend_preparation::Registration)->Self {Self::with_capabilities_identity(language,registration.identity())}
    fn with_capabilities_identity(language:Language,prepared:[usize;10])->Self {
        Self::with_options_identity(language, prepared, vybe_runtime::registry::compiler_optimizations(language))
    }
    fn with_options_identity(language:Language,prepared:[usize;10],options:vybe_runtime::registry::CompilerOptimizations)->Self {let hook=vybe_runtime::registry::hooks(language.name).compiler_optimizations;Self([
        language.parse as usize,language.profile_source as usize,
        language.emit_dispatch.map_or(0,|f|f as usize),language.normalize_class.map_or(0,|f|f as usize),
        language.register_tree.map_or(0,|f|f as usize),language.expand_source.map_or(0,|f|f as usize),
        hook.map_or(0,|f|f as usize),options.symbol_membership_index as usize,options.user_namespace_cache as usize,
        options.completion_worker_lowering as usize,options.compact_source_lines as usize,
        options.array_call_tail_identity as usize,options.interpolation_call_tail_identity as usize,
        options.fused_closure_analysis as usize,
        prepared[0],prepared[1],prepared[2],prepared[3],prepared[4],prepared[5],prepared[6],prepared[7],prepared[8],prepared[9],language.name.as_ptr() as usize,language.name.len(),
    ])}
}
// Explicit request-local scheduling permission, independent of lookup options.
type CompletionLoweringHook=fn(Language)->bool;
fn registered_completion_lowering(language:Language)->bool {
    vybe_runtime::registry::compiler_optimizations(language).completion_worker_lowering
}
#[derive(Debug)]
pub struct CompilationDependencies {
    certificate: Arc<std::sync::Mutex<crate::compile_dependencies::Certificate>>,
    libraries: Vec<(String,String,bool)>,
    profile: &'static str,
    profile_source: fn()-> &'static str,
    language_name: &'static str,
    frontend: FrontendIdentity,
}
impl CompilationDependencies {
    fn valid(&self,vm:&VM)->bool {
        (self.profile_source)()==self.profile
        && vybe_runtime::registry::find(self.language_name).is_some_and(|l|FrontendIdentity::of(l)==self.frontend)
        && self.certificate.lock().is_ok_and(|c|c.validates(&vm.modules) && c.runtime_helpers_valid(vm))
        && self.libraries.iter().all(|(name,identity,expected)|library_installed(vm,name,identity)==*expected)
    }
}
#[derive(Debug,Clone)]
pub struct DynamicCompilation {
    pub compile_dependencies: Option<Arc<CompilationDependencies>>,
    /// Validated common-AST augmentation-source declarations needed by later
    /// compilation in the same request. Replayed only at actual execution.
    pub source_declarations: Arc<[Statement]>,
    pub required_libraries: Vec<crate::preparation::LibraryDependency>,
    pub chunks: Vec<Chunk>,
    pub host_imports: HostImportMetadata,
    pub component_source: bool,
    pub component_binary: Option<std::sync::Arc<vybe_ast::canon::ComponentBinary>>,
    pub entry_path: Option<PathBuf>,
    /// Pending pure-source preparation, owned by this compilation's request.
    pub prepared_sources: Option<Arc<crate::preparation::PreparedSources>>,
    /// What the entry module DECLARED about presenting a UI — see
    /// [`vybe_ast::Directives::app_shell`]. `None` states nothing, and the
    /// document answers instead.
    pub app_shell: Option<vybe_ast::AppShell>,
}

impl DynamicCompilation {
    // Diagnostics only: the acceptance path continues to use the existing
    // predicates. Names and categories identify stale work without logging values.
    fn dependency_failure_for_audit(&self, vm: &VM) -> Option<String> {
        if let Some(dependencies) = &self.compile_dependencies {
            if (dependencies.profile_source)() != dependencies.profile {
                return Some("profile-source-changed".into());
            }
            if !vybe_runtime::registry::find(dependencies.language_name)
                .is_some_and(|language| FrontendIdentity::of(language) == dependencies.frontend) {
                return Some("frontend-changed".into());
            }
            match dependencies.certificate.lock() {
                Ok(certificate) => {
                    if !certificate.runtime_helpers_valid(vm) {
                        return Some("installed-runtime-helper-changed".into());
                    }
                    if let Some(failure) = certificate.validation_failure(&vm.modules) {
                        return Some(format!("certificate:{failure:?}"));
                    }
                }
                Err(_) => return Some("certificate-lock-poisoned".into()),
            }
            for (name, identity, expected) in &dependencies.libraries {
                let installed = library_installed(vm, name, identity);
                if installed != *expected {
                    return Some(format!("selected-library-changed:{name}:expected={expected}:installed={installed}"));
                }
            }
        }
        self.required_libraries.iter().find(|library|
            !library_installed(vm, &library.name, &library.identity))
            .map(|library| format!("required-library-unavailable:{}", library.name))
    }

    pub fn libraries_available(&self, vm: &VM) -> bool {
        self.compile_dependencies.as_ref().is_none_or(|c|c.valid(vm)) && self.required_libraries.iter().all(|library| {
            library_installed(vm, &library.name, &library.identity)
        })
    }
}

pub struct RuntimeCompilerService<'vm> {
    vm: &'vm mut VM,
    caps: Capabilities,
    php_runtime: PhpIncludeRuntime,
    js_runtime: JsDynamicRuntime,
}

/// Somewhere to keep a runtime include's compiled output between runs.
///
/// A runtime `require` compiles the file while the program is RUNNING, so the
/// server's entry-level cache — which wraps `compile_bundle` and is finished
/// long before `run_compiled` starts — never sees it. Under `--serve` that is
/// every include of every request recompiled from source.
///
/// The storage deliberately lives on the far side of this trait, in the
/// server, for the same reason the entry cache does: anything held in the VM
/// dies at `reset_to`, and anything held in a `static` here would be shared by
/// every tenant of the process with no owner to scope or clear it.
///
/// `fingerprint` is what makes a hit LEGAL — see [`module_fingerprint`].
pub trait IncludeCompileCache: Send + Sync {
    /// Exclusive request-lifetime storage. Cross-request, debugger/reset and
    /// external caches retain the default false policy and all raw snapshots.
    fn request_local_only(&self) -> bool { false }
    /// Remove only this successfully installed snapshot. The cache must not
    /// remove another path/fingerprint; false means no matching stored graph.
    fn retire_request_snapshot(&self, _path:&Path, _fingerprint:u64) -> bool { false }

    fn get(&self, path: &Path, fingerprint: u64) -> Option<Result<DynamicCompilation, String>>;
    fn store(
        &self,
        path: &Path,
        fingerprint: u64,
        deps: Option<Vec<PathBuf>>,
        result: Result<&DynamicCompilation, &str>,
    );
    /// Store an immutable raw snapshot without consuming a prepared owner.
    /// Existing caches remain compatible through the default implementation.
    fn store_with_chunks(
        &self, path: &Path, fingerprint: u64, deps: Option<Vec<PathBuf>>,
        result: Result<(&DynamicCompilation, &[Chunk]), &str>,
    ) {
        match result {
            Ok((metadata, chunks)) => {
                let mut raw = metadata.clone();
                raw.chunks = chunks.to_vec();
                self.store(path, fingerprint, deps, Ok(&raw));
            }
            Err(error) => self.store(path, fingerprint, deps, Err(error)),
        }
    }
}

/// Everything a runtime include's compile depends on BESIDES its own source.
///
/// Runtime includes are compiled as relocatable chunks. Calls into already
/// loaded code are still represented by import names and are resolved against
/// the live VM when the include runs, not baked as absolute function indices in
/// the cached chunk. That means the cache key must not include the whole VM
/// module table: WordPress grows that table on every include, so a plain
/// `require version.php` after early bootstrap and the same `require
/// version.php` while building the DB error page looked like different
/// compilations and missed the cache.
///
/// The source dependency set recorded by `record_source_reads` guards the file
/// contents. The remaining compile-time input that is intentionally observable
/// is the entry path: PHP magic constants and relative include resolution are
/// anchored to that top-level script.
pub fn module_fingerprint(vm: &VM, entry_path: &Path) -> u64 {
    use std::hash::{Hash, Hasher};

    fn hash_of(value: impl Hash) -> u64 {
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        value.hash(&mut hasher);
        hasher.finish()
    }

    let _ = vm;
    hash_of(entry_path)
}

pub fn run_with_js_dynamic_runtime(
    vm: &mut VM,
    caps: Capabilities,
    chunks: Vec<Chunk>,
) -> Result<Value, String> {
    ensure_js_runtime_registered(vm);
    let mut js_runtime = JsDynamicRuntime::new(caps);
    let _guard = js_runtime.activate(vm);
    vm.run(chunks).map_err(|err| err.to_string())
}

/// Install the compiler-as-a-service host imports used by runtime dynamic code
/// (`ecma:global.eval`, JS `Function`, and the language-generic
/// `vybe:eval.eval` used by PHP and Python).
///
/// Normal one-shot execution gets this through [`RuntimeCompilerService`].
/// Warm embedders call it during boot so the host module records are part of
/// the reset baseline instead of being dropped as per-tenant script state.
pub fn register_dynamic_runtime_imports(vm: &mut VM) {
    ensure_php_runtime_registered(vm);
    ensure_js_runtime_registered(vm);
}

struct PhpIncludeRuntime {
    caps: Capabilities,
    helper_session: crate::primitives::helper_templates::Session,
    #[cfg(any(test,v6_service_gate))]
    completion_handoff_hits: usize,
    #[cfg(any(test,v6_service_gate))]
    current_source_rank_receipts: usize,
    #[cfg(v6_service_gate)]
    committed_foreground_receipts: usize,
    #[cfg(v6_service_gate)]
    committed_foreground_ready_paths: usize,
    vm: *mut VM,
    current_paths: Vec<PathBuf>,
    /// Per-RUN, and it stays that way. Only the COMPILATION is cacheable: share
    /// this across requests and request two sees request one's "already
    /// included" and skips the include entirely.
    /// Canonical dynamic-load resources seen by PHP includes in this run.
    ///
    /// PHP's `_once` table is not populated only by `_once` calls: a plain
    /// `require 'autoload.php'` also makes a nested
    /// `require_once 'autoload.php'` a no-op while that file is still loading
    /// (and after it finishes). WordPress' sodium compat loader relies on that:
    /// `autoload.php` plain-requires `lib/php72compat.php`, whose first line is
    /// `require_once dirname(dirname(__FILE__)) . '/autoload.php'`.
    ///
    /// Plain `include` / `require` do not CONSULT this table for their own
    /// skip decision, so they can still intentionally load a file again later.
    /// They do PUBLISH loading/loaded state so subsequent `_once` calls observe
    /// PHP's included-files semantics.
    once_resources: HashMap<PathBuf, IncludeOnceState>,
    include_cache: Option<std::sync::Arc<dyn IncludeCompileCache>>,
    include_scope: Option<Value>,
    resident_units: HashMap<PathBuf,Arc<ResidentEntry>>,
    #[cfg(feature="service-verification")] resident_hits:usize,
    preparing: HashMap<PathBuf, crate::preparation::PreparationTicket>,
    prepared_ready: HashMap<PathBuf,crate::preparation::PreparedFile>,
    lowering: HashMap<PathBuf,LoweringTicket>,
    retired_compiled: RetiredCompiledBank,
    link_budget: Arc<AtomicUsize>,
    contextual_compile_failure: bool,
    fused_context:Option<Arc<FusedLoweringContext>>,
    completion_permission:Option<(FrontendIdentity,CompletionLoweringHook)>,
    completion_permission_override:bool,
    preparation_plan: crate::preparation::PreparationPlan,
    preparation_hints: crate::preparation::DeferredHints,
    builtin_callables: HashMap<(String, usize), Value>,
    php_trait_declarations: HashMap<String, Statement>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum IncludeOnceState {
    Loading,
    Loaded,
}

struct ActivePhpRuntimeGuard {
    previous: Option<*mut PhpIncludeRuntime>,
}

/// Input names are unique VM global bindings. A missing exact identity can
// shadow a lowercase identity, so retain precisely those negative bindings.
fn snapshot_library_bindings<'a>(bindings: impl Iterator<Item = (&'a str, Option<String>)>)
    -> HashMap<String, Option<String>>
{
    let mut libraries = HashMap::new();
    let mut possible_shadows = Vec::new();
    let mut max_ascii_identity_bytes = 0;
    for (name, identity) in bindings {
        let ascii = name.is_ascii();
        if identity.is_some() && ascii { max_ascii_identity_bytes = max_ascii_identity_bytes.max(name.len()); }
        if identity.is_some() || !ascii {
            libraries.insert(name.to_owned(), identity);
        } else if name.bytes().any(|byte| byte.is_ascii_uppercase()) {
            // Non-ASCII negative bindings retain the original representation above.
            // Already lowercase ASCII names cannot shadow a different key.
            // Keep borrowed names until positive identities are all collected.
            possible_shadows.push(name);
        }
    }
    if libraries.is_empty() { return libraries; }
    for name in possible_shadows {
        // ASCII lowercase preserves byte length. Unrelated large names cannot
        // shadow any retained ASCII library identity, so avoid copying them.
        if name.len() <= max_ascii_identity_bytes
            && libraries.get(&name.to_lowercase()).is_some_and(Option::is_some) {
            libraries.insert(name.to_owned(), None);
        }
    }
    libraries
}

// Request-local speculative result. Never enters the shared include cache
/// until the actual include validates its source and selected compile facts.
// Completion-worker preparation and lowering. The snapshot contains no VM
// pointer, request callbacks, mutable class objects or executable publication.
struct FusedLoweringContext {
    #[cfg(feature="service-verification")] manifest_probe:Option<Arc<FusedManifestProbe>>,
    modules:Arc<HashMap<String,vybe_runtime::ModuleRecord>>,
    libraries:HashMap<String,Option<String>>,
    profile:Arc<vybe_runtime::profile::LanguageProfile>,
    profile_source:&'static str,
    frontend:FrontendIdentity,
    admission:CompletionLoweringHook,
    helpers: Option<crate::primitives::helper_templates::Selected>,
    arguments: crate::primitives::argument_emission::Selection,
    runtime_helpers: std::sync::Mutex<Vec<vybe_runtime::shared_provider::InstalledRuntimeProvider>>,
}
impl FusedLoweringContext {
    fn capture(vm:&VM,language:Language,admission:CompletionLoweringHook, helpers:Option<crate::primitives::helper_templates::Selected>, registration:crate::frontend_preparation::Registration)->Result<Self,String>{
        let _cpu_scope = crate::compilation_timeline::span("worker_context_capture", Path::new(language.name));
        if FrontendIdentity::with_registration(language,registration)!=FrontendIdentity::of(language){return Err("frontend registration changed before context capture".into());}
        if helpers.as_ref().is_some_and(|selected| selected.selection.identity()!=registration.helper_selection().identity()){return Err("helper producer registration changed".into());}
        if !admission(language){return Err("frontend completion-lowering permission declined".into());}
        if crate::primitives::dump_classes_enabled()||["VYBE_TRACE","VYBE_DEBUG_IMPORTS","VYBE_SLOT_AUDIT"].iter().any(|k|std::env::var_os(k).is_some()){return Err("foreground diagnostics enabled".into());}
        let report = vm.debug_report();
        let _capture_phase = DebugPhase::new(report.clone(), "compiler worker context capture");
        let registration_phase = DebugPhase::new(report.clone(), "compiler worker context registration");
        let _=crate::primitives::instructions::host::CapabilityContext::get();
        crate::ensure_languages_registered();
        vybe_runtime::registry::register_all_platform_trees();vybe_runtime::registry::register_all_trees();
        drop(registration_phase);
        let profile_phase = DebugPhase::new(report.clone(), "compiler worker context profile");
        let source=(language.profile_source)();
        let profile=crate::bundle::cached_language_profile(source)?;
        drop(profile_phase);
        let libraries_phase = DebugPhase::new(report.clone(), "compiler worker context libraries");
        let libraries=snapshot_library_bindings(vm.global_index.keys().filter_map(|name|vm.global(name).map(|value|{
            let identity=match object_get_prop(value,crate::primitives::reflection::FIELD_LIBRARY_IDENTITY){Some(Value::String(id))=>Some(id.to_string()),_=>None};
            (name.as_str(),identity)
        })));
        drop(libraries_phase);
        let modules_phase = DebugPhase::new(report, "compiler worker context modules");
        let modules=Arc::new(vm.modules.iter().map(|(name,record)|{
            let mut copy=vybe_runtime::ModuleRecord::new_synthetic(name.clone());copy.exports=record.exports.clone();(name.clone(),copy)
        }).collect());
        drop(modules_phase);
        Ok(Self{
            #[cfg(feature="service-verification")] manifest_probe:FUSED_MANIFEST_PROBE.with(|slot|slot.borrow().clone()),
            modules,libraries,profile,profile_source:source,frontend:FrontendIdentity::with_registration(language,registration),admission,helpers,arguments:registration.argument_selection(),
            runtime_helpers:std::sync::Mutex::new(crate::primitives::bundle::snapshot_runtime_helpers(vm))})
    }
    fn installed(&self,name:&str,identity:&str)->bool{
        self.libraries.get(name).or_else(||self.libraries.get(&name.to_lowercase()))
            .is_some_and(|value|value.as_deref()==Some(identity))
    }
    fn lower_prepared(&self,prepared:&mut crate::preparation::PreparedFile,language:Language, link_budget: &Arc<AtomicUsize>,
    )->Option<LoweredSource>{
        let declined = if !prepared.matches_frontend(language) { Some("prepared-frontend-identity") }
            else if !prepared.dependencies.as_ref().is_some_and(Vec::is_empty) { Some("source-dependencies") }
            else if !prepared.module.as_ref().is_ok_and(independent_lowering) { Some("structural-admission") }
            else if FrontendIdentity::of(language)!=self.frontend { Some("frontend-identity") }
            else if !(self.admission)(language) { Some("frontend-permission") }
            else if (language.profile_source)()!=self.profile_source { Some("profile-source") }
            else { None };
        if let Some(reason)=declined {
            if std::env::var_os("VYBE_PREPARATION_AUDIT").is_some() {
                eprintln!("[preparation-audit] lowering-declined path={} reason={reason}",prepared.path.display());
            }
            return None;
        }
        let module=prepared.module.as_mut().unwrap();
        let mut libraries=Vec::new();let mut changed=Vec::new();
        fn bind(body:&mut[Statement],namespace:&str,path:&mut Vec<usize>,context:&FusedLoweringContext,
            decisions:&mut Vec<(String,String,bool)>,changed:&mut Vec<(Vec<usize>,vybe_ast::ClassLinkage)>,
        ){
            for (index,statement)in body.iter_mut().enumerate(){path.push(index);
                match &mut statement.kind{
                    StmtKind::ClassDecl{name,modifiers,..}=>{
                        if let vybe_ast::ClassLinkage::Library{identity}=&modifiers.linkage{
                            let name=if namespace.is_empty(){name.clone()}else{format!("{namespace}.{name}")};
                            let installed=context.installed(&name,identity);
                            decisions.push((name,identity.clone(),installed));
                            if installed{changed.push((path.clone(),modifiers.linkage.clone()));modifiers.linkage=vybe_ast::ClassLinkage::LibraryReference{identity:identity.clone()};}
                        }
                    },
                    StmtKind::NamespaceDecl{name,body}=>{
                        let name=if namespace.is_empty(){name.clone()}else{format!("{namespace}.{name}")};bind(body,&name,path,context,decisions,changed);
                    },
                    StmtKind::Block(body)=>bind(body,namespace,path,context,decisions,changed),
                    _=>{},
                };path.pop();
            }
        }
        bind(&mut module.body,"",&mut Vec::new(),self,&mut libraries,&mut changed);
        let started=std::time::Instant::now();
        let certificate=Arc::new(std::sync::Mutex::new(crate::compile_dependencies::Certificate::speculative()));
        certificate.lock().unwrap().install_runtime_helpers(self.runtime_helpers.lock().unwrap().clone());
        let compiled=std::panic::catch_unwind(std::panic::AssertUnwindSafe(||
            crate::bundle::record_source_reads(||if let Some(production)=prepared.statement_production.clone() {
                let use_borrowed=production.is_owned() && {
                    #[cfg(feature="service-verification")] {self.manifest_probe.as_ref().is_none_or(|probe|probe.use_borrowed())}
                    #[cfg(not(feature="service-verification"))] {true}
                };
                if use_borrowed{
                    Bundle::compile_source_statement_production_borrowed(language,&prepared.path,module,production,&self.modules,
                        Some(certificate.clone()),Some(self.profile.clone()),self.helpers.as_ref(),self.arguments)
                }else{
                    Bundle::compile_source_statement_production(language,&prepared.path,module.clone(),production,&self.modules,
                        Some(certificate.clone()),Some(self.profile.clone()),self.helpers.as_ref(),self.arguments)
                }
            }else{Bundle::compile_source_ast_certified(language,&prepared.path,module,&self.modules,Some(certificate.clone()),Some(self.profile.clone()),self.helpers.as_ref(),self.arguments)})));
        // Restore only exactly recorded declaration positions. Existing
        // LibraryReference nodes remain references. No Module/body clone.
        fn restore(body:&mut[Statement],path:&[usize],linkage:vybe_ast::ClassLinkage){
            let statement=&mut body[path[0]];
            if path.len()==1 {let StmtKind::ClassDecl{modifiers,..}=&mut statement.kind else{unreachable!()};modifiers.linkage=linkage;}
            else {match &mut statement.kind{StmtKind::NamespaceDecl{body,..}|StmtKind::Block(body)=>restore(body,&path[1..],linkage),_=>unreachable!()}}
        }
        for(path,linkage)in changed{restore(&mut module.body,&path,linkage);}
        let (compiled,reads)=match compiled{Ok(result)=>result,Err(_)=>return None};
        if !reads.as_ref().is_some_and(Vec::is_empty){
            certificate.lock().unwrap().reject();
            prepared.dependencies=reads; // Foreground cannot reuse an unvalidated source dependency.
            return None;
        }
        let compilation=compiled.map(|compiled|DynamicCompilation{
            compile_dependencies:Some(Arc::new(CompilationDependencies{certificate,libraries,profile:self.profile_source,profile_source:language.profile_source,language_name:language.name,frontend:self.frontend})),
            source_declarations:Arc::from([]),required_libraries:compiled.required_libraries,chunks:compiled.chunks,host_imports:compiled.host_imports,
            component_source:compiled.component_source,component_binary:compiled.component_binary,entry_path:Some(prepared.path.clone()),app_shell:compiled.app_shell,prepared_sources:None,
        });
        let compilation=compilation.map(|c| PendingCompilation::prepare(c, link_budget));
        #[cfg(feature="service-verification")]if let Some(probe)=&self.manifest_probe{probe.completed.fetch_add(1,Ordering::Relaxed); }
        Some(LoweredSource{elapsed:started.elapsed(),path:prepared.path.clone(),source:prepared.source.clone(),compilation})
    }
}

// Request-owned admission bounds retained speculative metadata. Rejection is
// always raw fallback; it does not block a worker or change include semantics.
const LINK_OWNER_BUDGET: usize = 64 * 1024 * 1024;
const LINK_REQUEST_BUDGET: usize = 128 * 1024 * 1024;
struct LinkPermit { budget: Arc<AtomicUsize>, bytes: usize }
impl LinkPermit {
    fn reserve(budget: &Arc<AtomicUsize>) -> Option<Self> {
        budget.fetch_update(Ordering::Relaxed, Ordering::Relaxed, |used|
            used.checked_add(LINK_OWNER_BUDGET).filter(|next| *next <= LINK_REQUEST_BUDGET)
        ).ok()?;
        Some(Self { budget: budget.clone(), bytes: LINK_OWNER_BUDGET })
    }
    fn retain(&mut self, bytes: usize) {
        debug_assert!(bytes <= self.bytes);
        self.budget.fetch_sub(self.bytes - bytes, Ordering::Relaxed);
        self.bytes = bytes;
    }
}
impl Drop for LinkPermit {
    fn drop(&mut self) { self.budget.fetch_sub(self.bytes, Ordering::Relaxed); }
}
enum LinkPayload {
    Raw(Vec<Chunk>),
    Prepared(vybe_runtime::PreparedLinkedCode, LinkPermit),
}
impl LinkPayload {
    fn chunks(&self) -> &[Chunk] {
        match self { Self::Raw(chunks) => chunks, Self::Prepared(code, _) => code.chunks() }
    }
}
// The public mutable compilation API stays raw. Only this private one-shot
// owner carries trusted metadata, with no mutable alias of its emitted chunks.
struct PendingCompilation { compilation: DynamicCompilation, code: LinkPayload }
impl PendingCompilation {
    fn raw(mut compilation: DynamicCompilation) -> Self {
        let chunks = std::mem::take(&mut compilation.chunks);
        Self { compilation, code: LinkPayload::Raw(chunks) }
    }
    fn prepare(compilation: DynamicCompilation, budget: &Arc<AtomicUsize>) -> Self {
        let mut pending = Self::raw(compilation);
        let Some(mut permit) = LinkPermit::reserve(budget) else { return pending; };
        let LinkPayload::Raw(chunks) = pending.code else { unreachable!() };
        let preparation_span = crate::compilation_timeline::span(
            "include_link_preparation", pending.compilation.entry_path.as_deref().unwrap_or(Path::new("")),
        );
        pending.code = match vybe_runtime::PreparedLinkedCode::prepare(chunks, LINK_OWNER_BUDGET) {
            Ok(code) => { permit.retain(code.metadata_bytes()); LinkPayload::Prepared(code, permit) }
            Err(chunks) => LinkPayload::Raw(chunks),
        };
        drop(preparation_span);
        pending
    }
    fn into_raw(mut self) -> DynamicCompilation {
        self.compilation.chunks = match self.code {
            LinkPayload::Raw(chunks) => chunks,
            LinkPayload::Prepared(code, _) => code.into_raw(),
        };
        self.compilation
    }
}

// Request-local reuse of successful completed work retired from lookahead.
// Bound is source+IR buffer capacities, NOT total retained heap or RSS.
const RETIRED_COMPILED_COUNT: usize = 8;
const RETIRED_COMPILED_PAYLOAD_BUDGET: usize = 16 * 1024 * 1024;
struct RetiredCompiledEntry { key: PathBuf, lowered: LoweredSource, charge: usize }
#[derive(Default)]
struct RetiredCompiledBank {
    entries: std::collections::VecDeque<RetiredCompiledEntry>,
    payload_bytes: usize,
    frontend: Option<FrontendIdentity>,
}
impl RetiredCompiledBank {
    fn clear(&mut self) { self.entries.clear();self.payload_bytes=0;self.frontend=None; }
    fn select(&mut self,frontend:Option<FrontendIdentity>) {
        if self.frontend!=frontend {self.clear();self.frontend=frontend;}
    }
    fn take(&mut self,key:&Path)->Option<LoweredSource> {
        let index=self.entries.iter().position(|entry|entry.key==key)?;
        let entry=self.entries.remove(index)?;
        self.payload_bytes-=entry.charge;
        Some(entry.lowered)
    }
    fn retain(&mut self,key:PathBuf,lowered:LoweredSource)->bool {
        let LoweredSource {elapsed,path,source,compilation}=lowered;
        let Ok(pending)=compilation else {return false};
        if self.frontend.is_none() || pending.compilation.compile_dependencies.as_ref()
            .is_none_or(|deps|Some(deps.frontend)!=self.frontend)
            || !pending.compilation.source_declarations.is_empty()
            || pending.compilation.prepared_sources.is_some()
            || pending.compilation.component_source
            || pending.compilation.component_binary.is_some() {return false;}
        // Move chunks out, destroy trusted metadata and release LinkPermit BEFORE
        // retaining/charging. No code/source/certificate clone or guest execution.
        let raw=pending.into_raw();
        let Some(charge)=raw.chunks.iter().try_fold(source.capacity(),|sum,chunk|
            sum.checked_add(chunk.code.capacity())) else {return false};
        if charge>RETIRED_COMPILED_PAYLOAD_BUDGET {return false;}
        drop(self.take(&key));
        while self.entries.len()>=RETIRED_COMPILED_COUNT
            || self.payload_bytes.checked_add(charge).is_none_or(|sum|sum>RETIRED_COMPILED_PAYLOAD_BUDGET) {
            let Some(entry)=self.entries.pop_front() else {return false};
            self.payload_bytes-=entry.charge;
        }
        self.payload_bytes+=charge;
        self.entries.push_back(RetiredCompiledEntry {key,charge,lowered:LoweredSource {
            elapsed,path,source,compilation:Ok(PendingCompilation::raw(raw)),
        }});
        true
    }
}

struct LoweredSource {
    elapsed: std::time::Duration,
    path: PathBuf,
    source: String,
    compilation: Result<PendingCompilation,String>,
}
pub(crate) struct LoweringTicket {
    link_budget: Arc<AtomicUsize>,
    result: crossbeam_channel::Receiver<LoweredSource>,
    cancelled: Arc<std::sync::atomic::AtomicBool>,
}
impl Drop for LoweringTicket {
    fn drop(&mut self) { self.cancelled.store(true,std::sync::atomic::Ordering::Relaxed); }
}
impl LoweredSource {
    fn valid(&self,bundle:&Bundle,vm:&VM)->bool {
        let source_matches=bundle.sources.first().is_some_and(|s|s.path==self.path && s.code==self.source);
        let dependencies_match=source_matches && self.compilation.as_ref().is_ok_and(|c|c.compilation.libraries_available(vm));
        if (!source_matches || !dependencies_match) && std::env::var_os("VYBE_PREPARATION_AUDIT").is_some() {
            let reason = if !source_matches {
                match bundle.sources.first() {
                    None => "source-missing".to_string(),
                    Some(source) if source.path != self.path => "source-path-changed".to_string(),
                    Some(_) => "source-text-changed".to_string(),
                }
            } else {
                match &self.compilation {
                    Err(_) => "worker-compilation-error".to_string(),
                    Ok(pending) => pending.compilation.dependency_failure_for_audit(vm)
                        .unwrap_or_else(|| "unclassified-dependency-failure".to_string()),
                }
            };
            eprintln!("[preparation-audit] lowering-rejected path={} source_equal={} dependencies_equal={} compile_ok={} reason={reason:?}",
                self.path.display(),source_matches,dependencies_match,self.compilation.is_ok());
        }
        source_matches && dependencies_match
    }
}
/// Pure scheduling traversal. Opted-in lookahead excludes deferred declaration
/// bodies at namespace/block boundaries; other statement forms retain the full
/// conservative walk. A skipped hint never suppresses an actual include.
fn walk_include_lookahead_exprs(statement: &Statement, include_deferred: bool,
    visit: &mut dyn FnMut(&vybe_ast::Expression),
) {
    if include_deferred { statement.walk_exprs(visit); return; }
    match &statement.kind {
        StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. }
        | StmtKind::StructDecl { .. } | StmtKind::ModuleDecl { .. } => {},
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
            for statement in body { walk_include_lookahead_exprs(statement, false, visit); }
        },
        _ => statement.walk_exprs(visit),
    }
}

/// Conservative structural eligibility. Unknown statement forms fall back;
/// statement-bearing expressions are checked as well as ordinary bodies.
pub(crate) fn independent_lowering(module:&Module)->bool {
    fn body(stmts:&[Statement])->bool {stmts.iter().all(stmt)}
    fn stmt(s:&Statement)->bool {
        let structural=match &s.kind {
            StmtKind::FunctionDecl{body:b,..}|StmtKind::NamespaceDecl{body:b,..}|StmtKind::Block(b)|StmtKind::DoWhile{body:b,..}=>body(b),
            StmtKind::If{then_body,elifs,else_body,..}=>body(then_body)&&elifs.iter().all(|(_,b)|body(b))&&else_body.as_ref().is_none_or(|b|body(b)),
            StmtKind::For{init,body:b,..}=>init.as_ref().is_none_or(|s|stmt(s))&&body(b),
            StmtKind::ForIn{body:b,else_body,..}|StmtKind::While{body:b,else_body,..}=>body(b)&&else_body.as_ref().is_none_or(|b|body(b)),
            StmtKind::Switch{cases,default,..}=>cases.iter().all(|c|body(&c.body))&&default.as_ref().is_none_or(|b|body(b)),
            StmtKind::Try{body:b,catches,else_body,finally}=>body(b)&&catches.iter().all(|c|body(&c.body))&&else_body.as_ref().is_none_or(|b|body(b))&&finally.as_ref().is_none_or(|b|body(b)),
            StmtKind::ClassDecl{members,modifiers,..}=> !matches!(modifiers.kind,vybe_ast::ClassKind::Trait|vybe_ast::ClassKind::Mixin|vybe_ast::ClassKind::Module)&&members.iter().all(|m|match m {
                ClassMember::Augment(_)=>false,
                ClassMember::Method(s)|ClassMember::NestedType(s)=>stmt(s),
                ClassMember::Constructor{body:b,..}=>body(b),
                ClassMember::Property{getter,setter,..}=>getter.as_ref().is_none_or(|b|body(b))&&setter.as_ref().is_none_or(|s|body(&s.body)),
                _=>true,
            }),
            StmtKind::Labeled{body:b,..}=>stmt(b),
            StmtKind::Empty|StmtKind::ScopeDecl{..}|StmtKind::Echo(_)|StmtKind::Exit{..}|StmtKind::Delete(_)|StmtKind::Assert{..}|StmtKind::GoTo(_)|StmtKind::Label(_)|StmtKind::Expr(_)|StmtKind::Directive{..}|StmtKind::VarDecl{..}|StmtKind::InterfaceDecl{..}|StmtKind::Return(_)|StmtKind::Break(_)|StmtKind::Continue(_)|StmtKind::Throw{..}|StmtKind::Assign{..}|StmtKind::CompoundAssign{..}=>true,
            _=>false,
        };
        structural
    }
    if !module.imports.iter().all(|i|matches!(i.kind,vybe_ast::ImportKind::Simple{..}|vybe_ast::ImportKind::Wildcard{..})) || !body(&module.body) {return false;}
    let mut expressions=true;
    for s in &module.body {s.walk_exprs(&mut |e|match &e.kind {
        vybe_ast::ExprKind::ClassExpr{..}=>expressions=false,
        vybe_ast::ExprKind::FunctionExpr(s)=>expressions &= stmt(s),
        vybe_ast::ExprKind::Lambda{body:vybe_ast::LambdaBody::Block(b),..}=>expressions &= body(b),
        _=>{},
    });}
    expressions
}

struct JsDynamicRuntime {
    caps: Capabilities,
    vm: *mut VM,
}

struct ConstructedJsFunction {
    caps: Capabilities,
    vm: VM,
    function: Value,
}

struct ActiveJsRuntimeGuard {
    previous: Option<*mut JsDynamicRuntime>,
}

impl<'vm> RuntimeCompilerService<'vm> {
    pub fn new(vm: &'vm mut VM) -> Self {
        Self::with_capabilities(vm, Capabilities::all())
    }

    pub fn with_capabilities(vm: &'vm mut VM, caps: Capabilities) -> Self {
        Self {
            vm,
            caps: caps.clone(),
            php_runtime: PhpIncludeRuntime::new(caps.clone()),
            js_runtime: JsDynamicRuntime::new(caps),
        }
    }




    pub fn vm(&mut self) -> &mut VM {
        self.vm
    }

    /// Select request-owned helper reuse explicitly. None remains disabled
    /// across compiler-service handoffs; ordinary unregistered emitters are unchanged.
    pub fn set_helper_pool(&mut self, pool: Option<Arc<crate::primitives::helper_templates::Pool>>) {
        self.php_runtime.helper_session.set_pool(pool);
        self.php_runtime.fused_context=None;
        self.php_runtime.lowering.clear();self.php_runtime.retired_compiled.clear();
    }

    /// Give runtime includes somewhere to reuse a compilation from. Without
    /// one they compile from source every time, which is the behaviour this
    /// service has always had.
    pub fn set_include_cache(&mut self, cache: std::sync::Arc<dyn IncludeCompileCache>) {
        self.php_runtime.include_cache = Some(cache);
    }

    /// Opt in to completion-worker lowering for one exact registered frontend.
    /// The caller supplies an admission hook independently from compiler lookup
    /// acceleration. None disables scheduling and cancels pending lower results;
    /// prepared source tickets remain available for foreground compilation.
    pub fn set_completion_lowering(&mut self,frontend:Language,hook:Option<fn(Language)->bool>)->Result<(),String>{
        if hook.is_none(){self.php_runtime.preparation_plan.set_retirement_admission(false);self.php_runtime.lowering.clear();self.php_runtime.retired_compiled.clear();self.php_runtime.fused_context=None;self.php_runtime.completion_permission=None;self.php_runtime.completion_permission_override=true;return Ok(());}
        let registered=crate::languages::find_by_name(frontend.name).ok_or("unregistered completion frontend")?;
        if FrontendIdentity::of(registered)!=FrontendIdentity::of(frontend){return Err("completion frontend differs from registered parser/hooks".into());}
        self.php_runtime.lowering.clear();self.php_runtime.retired_compiled.clear();self.php_runtime.fused_context=None;
        self.php_runtime.completion_permission=hook.map(|hook|(FrontendIdentity::of(frontend),hook));
        self.php_runtime.completion_permission_override=true;
        Ok(())
    }

    pub fn compile_bundle(&mut self, bundle: &Bundle) -> Result<DynamicCompilation, String> {
        self.compile_bundle_with_end_hooks(bundle, true)
    }

    fn compile_bundle_with_end_hooks(&mut self, bundle: &Bundle,
        run_module_end_hooks: bool) -> Result<DynamicCompilation, String> {
        let _cpu_scope = crate::compilation_timeline::span("entry_compile_service",
            bundle.sources.first().map_or(Path::new("<bundle>"), |source| source.path.as_path()));
        let registration=crate::frontend_preparation::snapshot(bundle.language.parse);
        let helpers=self.php_runtime.helper_session.selected(registration.helper_selection());
        if registration.capabilities.prepare.is_some() && std::env::var_os("VYBE_PREPARATION_AUDIT").is_some() {
            eprintln!("[preparation-audit] frontend-source-unit-carrier-v1 parser={} early_scout={}",bundle.language.name,registration.capabilities.early_hints);
        }
        self.php_runtime.select_completion_permission(bundle.language);
        if self.php_runtime.completion_permission.is_some_and(|(identity,hook)|identity==FrontendIdentity::of(bundle.language)&&hook(bundle.language)) {
            self.php_runtime.fused_context=FusedLoweringContext::capture(self.vm,bundle.language,self.php_runtime.completion_permission.unwrap().1,helpers.clone(),registration).ok().map(Arc::new);
        }else{self.php_runtime.fused_context=None;}
        let _report_scope = DebugReportScope::enter(self.vm.debug_report());
        let _phase = DebugPhase::new(self.vm.debug_report(), "compile entry");
        ensure_php_runtime_registered(self.vm);
        ensure_js_runtime_registered(self.vm);
        if bundle.language.name == "php" && (registration.capabilities.prepare.is_none() || registration.capabilities.early_hints) {
            if let Some(source) = bundle.sources.first() {
                self.php_runtime.prepare_php_source_hints(bundle, &source.path, Some(self.vm));
            }
        }
        let preparer=registration.preparer();
        let prepared_result=bundle.prepared_frontend_with_sink(None,preparer,&mut |hints| {
            if let Some(source)=bundle.sources.first(){self.php_runtime.consume_current_source_preparation_hints(source.path.clone(),hints,&source.path,bundle.language,self.vm);}
        });
        let mut prepared=match prepared_result {Ok(result)=>result,Err(error)=>{self.php_runtime.cancel_request_preparation();return Err(error)}};
        let preparation_hints_complete=prepared.preparation_hints_complete;
        if let (Some(hints),Some(source))=(prepared.hints.take(),bundle.sources.first()) {
            self.php_runtime.consume_current_source_preparation_hints(source.path.clone(),hints,&source.path,bundle.language,self.vm);
        }
        let statement_production=prepared.statement_production.take();
        let mut module=prepared.module;
        if bundle.language.name == "php" && !preparation_hints_complete {
            if let Some(source) = bundle.sources.first() {
                self.php_runtime.prepare_known_php_includes(&module, bundle, &source.path);
            }
        }
        let source_declarations = crate::preparation::source_declarations(&module);
        bind_installed_libraries(&mut module, self.vm);
        let result=if let Some(production)=statement_production {
            bundle.compile_prepared_statement_production(module,production,&self.vm.modules,run_module_end_hooks,None,None,helpers.as_ref(),registration.argument_selection())
        }else{bundle.compile_prepared_certified_with_emitters(&module,&self.vm.modules,run_module_end_hooks,None,None,helpers.as_ref(),registration.argument_selection())};
        let compiled=match result{Ok(compiled)=>compiled,Err(error)=>{self.php_runtime.cancel_request_preparation();return Err(error)}};
        let pending = std::mem::take(&mut self.php_runtime.preparing);
        let plan = std::mem::take(&mut self.php_runtime.preparation_plan);
        let deferred = std::mem::take(&mut self.php_runtime.preparation_hints);
        let lowerings = std::mem::take(&mut self.php_runtime.lowering);
        let prepared_sources = (!pending.is_empty() || !plan.is_empty() || !deferred.is_empty() || !lowerings.is_empty() || self.php_runtime.helper_session.has_request_choice()).then(||
            Arc::new(crate::preparation::PreparedSources::new(pending, plan, deferred)
                .with_lowerings(bundle.language, lowerings).with_helper_session(self.php_runtime.helper_session.clone())));
        Ok(DynamicCompilation {
            compile_dependencies: None,
            source_declarations: source_declarations.into(),
            required_libraries: compiled.required_libraries,
            chunks: compiled.chunks,
            host_imports: compiled.host_imports,
            component_source: compiled.component_source,
            component_binary: compiled.component_binary,
            entry_path: bundle.sources.first().map(|source| source.path.clone()),
            app_shell: compiled.app_shell,
            prepared_sources,
        })
    }

    pub fn compile_path(&mut self, path: &Path) -> Result<DynamicCompilation, String> {
        let bundle = crate::projects::load(path)?;
        self.compile_bundle(&bundle)
    }

    pub fn compile_source(
        &mut self,
        source: impl Into<String>,
        language: Language,
        virtual_path: impl Into<PathBuf>,
    ) -> Result<DynamicCompilation, String> {
        self.ensure_dynamic_compile_allowed()?;
        let bundle = bundle_from_source(source, language, virtual_path);
        self.compile_bundle(&bundle)
    }

    pub fn compile_source_by_name(
        &mut self,
        source: impl Into<String>,
        language_name: &str,
        virtual_path: impl Into<PathBuf>,
    ) -> Result<DynamicCompilation, String> {
        let language = languages::find_by_name(language_name)
            .ok_or_else(|| format!("unknown language: {language_name}"))?;
        self.compile_source(source, language, virtual_path)
    }

    /// Compile one complete interactive submission against the live session.
    /// Only the new source is parsed/emitted. Previously installed chunks stay
    /// resident; normal registered frontend hooks and library binding apply.
    /// End-of-program hooks are deliberately deferred: a submission is not
    /// the end of the session. This accepts complete syntax units, not partial
    /// token streams, and does not retain the submission's AST after emission.
    pub fn compile_submission_source(
        &mut self,
        source: impl Into<String>,
        language: Language,
        virtual_path: impl Into<PathBuf>,
    ) -> Result<DynamicCompilation, String> {
        self.ensure_dynamic_compile_allowed()?;
        self.compile_bundle_with_end_hooks(&bundle_from_source(source, language, virtual_path), false)
    }

    /// Install and execute a compiled submission by appending its code to the
    /// existing VM. Nested execution preserves the caller's execution context
    /// and restores its import tables, including on an execution error.
    pub fn run_submission(&mut self, compiled: DynamicCompilation) -> Result<Value, String> {
        self.run_compiled_impl(compiled, true)
    }

    /// Compile and execute a complete incremental unit without ending the
    /// session or recompiling code installed by earlier submissions.
    pub fn submit_source(
        &mut self,
        source: impl Into<String>,
        language: Language,
        virtual_path: impl Into<PathBuf>,
    ) -> Result<Value, String> {
        let compiled = self.compile_submission_source(source, language, virtual_path)?;
        self.run_submission(compiled)
    }

    pub fn can_dynamic_compile(&self) -> bool {
        self.caps.has(Capability::DynamicCompile)
    }

    fn ensure_dynamic_compile_allowed(&self) -> Result<(), String> {
        if self.can_dynamic_compile() {
            return Ok(());
        }
        Err("Dynamic compilation is disabled by the current capability set (missing Capability::DynamicCompile)".to_string())
    }

    pub fn run_compiled(&mut self, compiled: DynamicCompilation) -> Result<Value, String> {
        self.run_compiled_impl(compiled, false)
    }

    /// Compile and run one non-entry unit of a multi-language program.
    ///
    /// This is the linking step. Each language is compiled on its own terms —
    /// its own grammar, profile and prelude — and then loaded into the *same*
    /// VM, so its functions and classes land in the shared global table and its
    /// top-level code runs before the entry unit does. Nothing is concatenated
    /// across languages; the chunk index, import-table and type-table
    /// relocation that makes this safe is `VM::run_linked_nested`, the same
    /// machinery `eval` uses to define real globals in a live VM.
    ///
    /// `nested` matters: the unit runs under *its own* resolved import table,
    /// which is restored afterwards so the entry unit's table is the one active
    /// when it starts.
    pub fn run_program_unit(&mut self, bundle: &Bundle) -> Result<Value, String> {
        let compiled = self.compile_bundle(bundle)?;
        self.run_compiled_impl(compiled, true)
    }

    fn run_compiled_impl(
        &mut self,
        compiled: DynamicCompilation,
        nested: bool,
    ) -> Result<Value, String> {
        if !compiled.libraries_available(self.vm) {
            return Err("compiled unit requires library definitions that are not installed in this VM".into());
        }
        let mut deferred = Default::default();
        let mut request_lowerings=HashMap::new();
        if let Some(prepared) = &compiled.prepared_sources {
            let batch = prepared.take();
            self.php_runtime.helper_session.adopt(&batch.helper_session);
            if let Some(frontend) = batch.lowering_frontend {
                self.php_runtime.select_completion_permission(frontend);
                let permitted = self.php_runtime.completion_permission.is_some_and(|(identity, hook)|
                    identity == FrontendIdentity::of(frontend) && hook(frontend));
                if permitted {
                    // Adopt the same request admission pool as pending jobs;
                    // creating a fresh service must not reset their bound.
                    if let Some(ticket) = batch.lowerings.values().next() {
                        self.php_runtime.link_budget = ticket.link_budget.clone();
                    }
                    request_lowerings.extend(batch.lowerings);
                    let registration=crate::frontend_preparation::snapshot(frontend.parse);
                    let helpers=self.php_runtime.helper_session.selected(registration.helper_selection());
                    self.php_runtime.fused_context = FusedLoweringContext::capture(
                        self.vm, frontend, self.php_runtime.completion_permission.unwrap().1,
                        helpers,registration,
                    ).ok().map(Arc::new);
                } else {
                    self.php_runtime.fused_context = None;
                    self.php_runtime.lowering.clear();self.php_runtime.retired_compiled.clear();
                    drop(batch.lowerings);
                }
                // Dropping unpermitted receivers cancels only speculation;
                // their original parse tickets remain available below.
            }
            deferred = batch.deferred;
            self.php_runtime.preparation_plan.extend(batch.plan);
            for (path, ticket) in batch.tickets {
                if let Some(lowering)=self.php_runtime.lowering.remove(&path){request_lowerings.insert(path.clone(),lowering);}
                if self.php_runtime.preparing.len() < 4 {
                    self.php_runtime.preparing.entry(path).or_insert(ticket);
                }
            }
            request_lowerings.retain(|path, _| self.php_runtime.preparing.contains_key(path));
            self.php_runtime.replenish_php_preparation();
        }
        // Replenishment may create new fused jobs after taking the batch.
        // Keep only receivers backed by this request's preparation tickets.
        {let preparing=&self.php_runtime.preparing;
         request_lowerings.extend(self.php_runtime.lowering.drain().filter(|(path,_)|preparing.contains_key(path)));}
        ensure_php_runtime_registered(self.vm);
        ensure_js_runtime_registered(self.vm);
        let base_chunk_index = self.vm.chunks.len();
        crate::host_imports::install(self.vm, &compiled.host_imports);
        install_chunk_globals(self.vm, &compiled.chunks, base_chunk_index);
        let active_imports = compiled
            .chunks
            .first()
            .map(|chunk| chunk.imports.as_slice())
            .unwrap_or(&[]);
        let active_resolved_imports = resolve_imports(self.vm, active_imports)?;
        let _php_runtime = self.php_runtime.activate(
            self.vm,
            compiled.entry_path.as_deref(),
        );
        self.php_runtime.lowering.extend(request_lowerings);
        self.php_runtime.preparation_hints = deferred;
        self.php_runtime.install_source_declarations(&compiled.source_declarations);
        let _js_runtime =
            self.js_runtime
                .activate(self.vm);
        if nested {
            return self
                .vm
                .run_linked_nested(compiled.chunks, active_resolved_imports)
                .map_err(|e| e.to_string());
        }
        self.vm.run(compiled.chunks).map_err(|e| e.to_string())
    }

    pub fn compile_and_run_bundle(&mut self, bundle: &Bundle) -> Result<Value, String> {
        let compiled = self.compile_bundle(bundle)?;
        self.run_compiled(compiled)
    }

    pub fn compile_and_run_path(&mut self, path: &Path) -> Result<Value, String> {
        let compiled = self.compile_path(path)?;
        self.run_compiled(compiled)
    }

    pub fn compile_and_run_source(
        &mut self,
        source: impl Into<String>,
        language: Language,
        virtual_path: impl Into<PathBuf>,
    ) -> Result<Value, String> {
        let compiled = self.compile_source(source, language, virtual_path)?;
        self.run_compiled(compiled)
    }
}

pub fn bundle_from_source(
    source: impl Into<String>,
    language: Language,
    virtual_path: impl Into<PathBuf>,
) -> Bundle {
    let path = virtual_path.into();
    // A physical PHP source has one identity across include preparation and
    // foreground compilation, including symlink aliases such as /var on macOS.
    // Resolve before expanding magic constants, as PHP does for __FILE__/__DIR__.
    // Eval and other virtual/nonexistent paths retain their supplied spelling.
    let path = if language.name == "php" {
        path.canonicalize().unwrap_or(path)
    } else {
        path
    };
    let _cpu_scope = crate::compilation_timeline::span("source_bundle", &path);
    let source = source.into();
    let name = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .unwrap_or("dynamic")
        .to_string();
    let code = if language.name == "php" {
        preprocess_dynamic_php_source(&source, &path)
    } else {
        source
    };

    Bundle {
        name,
        language,
        sources: vec![SourceFile { path, code }],
        wasm_files: Vec::new(),
        entry_point: EntryPoint::Auto,
    }
}

#[cfg(all(test, unix))]
#[path = "dynamic_source_identity_tests.rs"]
mod php_source_identity_tests;

pub fn language_for_path(path: &Path) -> Option<Language> {
    let ext = path.extension()?.to_str()?;
    languages::find_by_extension(ext)
}

/// Debugger expression evaluator — LANGUAGE-AGNOSTIC (common AST + common
/// compiler). Given the live paused VM, an expression written in the program's
/// language, and the paused frame's locals, evaluate it in an ISOLATED mini-VM
/// (the live VM's stack/frames are never touched) and return the value.
///
/// How it stays general across every language (JS, Python, PHP, …):
///   1. Parse the expression with the language's own `parse` — which, like every
///      Vybe frontend, targets the SAME common AST (`vybe_ast::Module`).
///   2. Rewrite the trailing expression statement into a `Return` at the AST
///      level, so the program's result IS the expression's value. No per-language
///      source template — the value capture is a common-AST transform.
///   3. Compile with the language's profile via the COMMON compiler.
///   4. Run; a top-level `return` yields the value from `run()`.
/// Locals + copyable globals are injected as the mini-VM's globals so names
/// resolve exactly as the program's language would.
pub fn debug_eval_expression(
    live: &VM,
    expr: &str,
    locals: &[(String, Value)],
    language: Language,
    caps: Capabilities,
) -> Result<Value, String> {
    let expr = expr.trim().trim_end_matches(';');
    if expr.is_empty() {
        return Err("empty expression".into());
    }

    // 1. Parse the expression in the program's language → common AST, then hold
    //    onto both the prelude-carrying module and the extracted `Expression`.
    //    Dynamic languages parse a bare expression directly; statically-structured
    //    languages (Go/Java/C#/C/VB) need minimal scaffolding to parse a fragment,
    //    from which we lift the same common-AST expression node.
    let (bundle, mut module, extracted) = obtain_eval_expression(language, expr)?;

    // 2. Common-AST value capture: replace the trailing expression statement with
    //    a function that takes the paused frame's locals as PARAMETERS and
    //    `return`s the expression. Passing locals as parameters (not injected
    //    globals) is what makes this work for EVERY language: a parameter is in
    //    scope inside the function body regardless of the language's global-scope
    //    rules (JS/Python fall back to globals for free vars, PHP does not — but
    //    all languages see their parameters). The node is a common `FunctionDecl`,
    //    built once for all languages; the values are passed at invoke time.
    const EVAL_FN: &str = "__vybe_dbg_eval__";
    // One parameter per copyable local (its compiler-scope name, e.g. `$x` in
    // PHP, `x` in JS/Python), invoked with the local's live value.
    let frame_args: Vec<(String, Value)> = locals
        .iter()
        .filter(|(_, v)| eval_value_is_copyable(v))
        .cloned()
        .collect();
    let params: Vec<crate::ast::Param> = frame_args
        .iter()
        .map(|(name, _)| crate::ast::Param {
            name: name.clone(),
            type_hint: None,
            default: None,
            pass_by: crate::ast::PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        })
        .collect();
    module.body.push(crate::ast::Statement::new(
        crate::ast::StmtKind::FunctionDecl {
            name: EVAL_FN.to_string(),
            params,
            return_type: None,
            body: vec![crate::ast::Statement::new(crate::ast::StmtKind::Return(
                Some(extracted),
            ))],
            modifiers: crate::ast::Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
    ));

    // 3. Set up the isolated mini-VM (runtime registration + host functions).
    //    Register the same platform exports before overlaying live host closures.
    let mut eval_vm = VM::new();
    crate::primitives::platforms::register_platforms(&mut eval_vm, &caps);
    ensure_php_runtime_registered(&mut eval_vm);
    ensure_js_runtime_registered(&mut eval_vm);
    // Share the LIVE program's host-function closures (matched by name), so host
    // calls in the eval expression hit the live captured state (e.g. the shared
    // GuiState Arc) instead of this mini-VM's fresh, empty one. Execution and
    // exception state stay isolated — a throw is contained to `eval_vm`, never
    // corrupting the paused program's handler stack. See `overlay_host_fns_from`.
    eval_vm.overlay_host_fns_from(live);
    // Copy live globals (scalars + shared objects; NOT function values — their
    // chunk_index refs belong to the live VM's chunk table). Objects are shared
    // by Arc, so reads see live object state. Locals come in as params (below).
    for (k, v) in live.globals_by_name() {
        if eval_value_is_copyable(&v) {
            eval_vm.set_global(&k, v);
        }
    }

    // 4. Compile the transformed module with the bundle's FULL pipeline, run it
    //    (defines the eval function as a global), then invoke it with the frame's
    //    local values — the value it returns is the expression's value.
    let compiled = bundle
        .compile_prepared_module(&module, &eval_vm.modules)
        .map_err(|e| format!("eval compile error: {e}"))?;
    let dynamic = into_dynamic_compilation(compiled);
    {
        let mut service = RuntimeCompilerService::with_capabilities(&mut eval_vm, caps);
        service
            .run_compiled(dynamic)
            .map_err(|e| format!("eval error: {e}"))?;
    }
    let fn_val = eval_vm
        .remove_global(EVAL_FN)
        .or_else(|| eval_vm.remove_global(&EVAL_FN.to_lowercase()))
        .ok_or("eval function did not compile")?;
    let arg_values: Vec<Value> = frame_args.into_iter().map(|(_, v)| v).collect();
    let result = eval_vm.invoke_callback(&fn_val, &arg_values);
    if let Some(exc) = eval_vm.last_exception.take() {
        return Err(format!("eval threw: {exc}"));
    }
    Ok(result)
}

/// Parse `expr` in `language` → `(bundle, prelude-carrying module, expression)`.
/// Dynamic / expression-oriented languages parse a bare expression directly (its
/// trailing statement is an `Expr` we lift out). Statically-structured languages
/// (Go/Java/C#/C/VB) get minimal scaffolding so the fragment parses at all, and
/// we lift the same common-AST expression from the wrapper's `return` — the
/// scaffold only makes it *parseable*; its operator semantics still come from the
/// common compiler.
fn obtain_eval_expression(
    language: Language,
    expr: &str,
) -> Result<(Bundle, crate::ast::Module, crate::ast::Expression), String> {
    use crate::ast::StmtKind;
    // Path A — bare expression (dynamic / expression-oriented languages).
    if language.name != "php" {
        let bundle = bundle_from_source(expr, language, PathBuf::from("<dbg-eval>"));
        if let Ok(mut module) = bundle.prepared_module() {
            if let Some(idx) = module
                .body
                .iter()
                .rposition(|s| matches!(s.kind, StmtKind::Expr(_)))
            {
                if let StmtKind::Expr(e) =
                    std::mem::replace(&mut module.body[idx].kind, StmtKind::Empty)
                {
                    return Ok((bundle, module, e));
                }
            }
        }
    }
    // Path B — scaffold a fragment for statically-structured languages, then lift
    // the expression from the wrapper function's `return` (the uncalled wrapper
    // is harmless — only our params-function is invoked).
    if let Some(scaffolded) = eval_scaffold(language.name, expr) {
        let bundle = bundle_from_source(scaffolded, language, PathBuf::from("<dbg-eval>"));
        let module = bundle
            .prepared_module()
            .map_err(|e| format!("parse error: {e}"))?;
        if let Some(e) = frag_return_expression(&module) {
            return Ok((bundle, module, e));
        }
    }
    Err(format!(
        "expression eval isn't available for '{}' yet — use `p <name>` for a read",
        language.name
    ))
}

/// Minimal per-language source that makes an expression parse as a program,
/// exposing it as `__vybe_frag`'s return. Lives in the eval-service layer (not
/// the emit compiler), the same layer that already frames per-language source
/// (e.g. PHP `<?php`). `None` for languages that already parse bare expressions
/// or aren't expression-evaluable here.
fn eval_scaffold(language_name: &str, expr: &str) -> Option<String> {
    let src = match language_name {
        "go" => format!(
            "package main\nfunc main() {{}}\nfunc __vybe_frag() interface{{}} {{\n\treturn ({expr})\n}}\n"
        ),
        "java" => {
            format!("class __VybeFrag {{ static Object __vybe_frag() {{ return ({expr}); }} }}\n")
        }
        "csharp" | "cs" => {
            format!("class __VybeFrag {{ static object __vybe_frag() {{ return ({expr}); }} }}\n")
        }
        "c" => format!("long __vybe_frag() {{ return ({expr}); }}\n"),
        "dart" => format!("dynamic __vybe_frag() {{ return ({expr}); }}\n"),
        "vb" => format!(
            "Module __VbFrag\n  Function __vybe_frag() As Object\n    Return ({expr})\n  End Function\nEnd Module\n"
        ),
        "php" => format!("<?php function __vybe_frag() {{ return ({expr}); }}\n"),
        _ => return None,
    };
    Some(src)
}

/// Lift the expression from `__vybe_frag`'s `return`, whether it's a top-level
/// function (Go/C) or a method inside a class/module (Java/C#/VB).
fn frag_return_expression(module: &crate::ast::Module) -> Option<crate::ast::Expression> {
    use crate::ast::{ClassMember, StmtKind};
    fn return_of(name: &str, body: &[crate::ast::Statement]) -> Option<crate::ast::Expression> {
        if name != "__vybe_frag" {
            return None;
        }
        body.iter().find_map(|bs| match &bs.kind {
            StmtKind::Return(Some(e)) => Some(e.clone()),
            _ => None,
        })
    }
    for s in &module.body {
        match &s.kind {
            StmtKind::FunctionDecl { name, body, .. } => {
                if let Some(e) = return_of(name, body) {
                    return Some(e);
                }
            }
            StmtKind::ClassDecl { members, .. } => {
                for m in members {
                    if let ClassMember::Method(stmt) = m {
                        if let StmtKind::FunctionDecl { name, body, .. } = &stmt.kind {
                            if let Some(e) = return_of(name, body) {
                                return Some(e);
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }
    None
}

/// A value that can be lifted into the eval mini-VM for inspection. Function
/// objects retain their type and properties; dispatch guards their VM-local
/// chunk indexes if an expression tries to call them.
fn eval_value_is_copyable(v: &Value) -> bool {
    // Preserve callable values for read-only expressions such as `typeof f`
    // and `C.prototype.method`. Their chunk indexes belong to the live VM;
    // attempting to invoke one in the isolated VM is guarded at dispatch.
    let _ = v;
    true
}

pub fn install_chunk_globals(vm: &mut VM, chunks: &[Chunk], base_chunk_index: usize) {
    install_chunk_globals_with_policy(vm, chunks, base_chunk_index, false);
}

fn install_chunk_globals_with_policy(
    vm: &mut VM,
    chunks: &[Chunk],
    base_chunk_index: usize,
    preserve_existing_globals: bool,
) {
    for (idx,chunk) in chunks.iter().enumerate() {
        install_callable_global(vm,&chunk.name,chunk.arity,base_chunk_index+idx,preserve_existing_globals);
    }
}
fn install_callable_global(vm:&mut VM,declared_name:&str,arity:u8,chunk_index:usize,preserve_existing_globals:bool) {
        if !should_publish_chunk_name(declared_name) {
            return;
        }
        let name = declared_name.to_lowercase();
        if preserve_existing_globals && vm.has_global(&name) {
            return;
        }
        // A dynamically included function may share its PHP name with an
        // already declared class. Keep the class constructor in the bare
        // global; function calls use their separate callable registration.
        let existing_type = vm
            .global(&name)
            .and_then(|value| object_get_prop(value, crate::primitives::reflection::FIELD_KIND))
            .is_some_and(|kind| {
                matches!(
                    kind,
                    Value::String(ref name)
                        if matches!(
                            name.as_ref(),
                            "class" | "interface" | "trait" | "mixin" | "module" | "struct"
                        )
                )
            });
        if existing_type {
            return;
        }

        let func = Function {
            name: Some(declared_name.to_string()),
            arity: arity,
            chunk_index: chunk_index,
            upvalues: vec![],
        };
        let mut obj = Object::new();
        obj.kind = ObjectKind::Function(func);
        let val = Value::Object(vybe_runtime::heap::alloc(obj));
        vm.set_global_owned(name, val);
}

fn should_publish_chunk_name(name: &str) -> bool {
    !name.is_empty()
        && name != "<script>"
        && name != "<bootstrap>"
        && !name.starts_with("__stdlib_")
}

fn preprocess_dynamic_php_source(source: &str, source_path: &Path) -> String {
    // The replacement scanner only recognizes these exact markers. In their
    // absence all bytes are retained, including comments and quoted strings.
    if !source.contains("__FILE__") && !source.contains("__DIR__") {
        return source.to_owned();
    }
    let absolute = absolutize_path(source_path);
    let file_literal = php_single_quoted_literal(&absolute.to_string_lossy());
    let dir_literal = php_single_quoted_literal(
        absolute
            .parent()
            .unwrap_or(Path::new("."))
            .to_string_lossy()
            .as_ref(),
    );

    replace_php_magic_constants(source, &file_literal, &dir_literal)
}

fn library_installed(vm: &VM, name: &str, identity: &str) -> bool {
    vm.global(name).or_else(|| vm.global(&name.to_lowercase()))
        .and_then(|value| object_get_prop(value,
            crate::primitives::reflection::FIELD_LIBRARY_IDENTITY))
        .is_some_and(|value| matches!(value, Value::String(ref current) if current.as_ref() == identity))
}

fn bind_installed_libraries(module: &mut vybe_ast::Module, vm: &VM) {
    crate::preparation::bind_installed_libraries(module, |name, identity| {
        library_installed(vm, name, identity)
    });
}

fn replace_php_magic_constants(source: &str, file_literal: &str, dir_literal: &str) -> String {
    #[derive(Clone, Copy, PartialEq, Eq)]
    enum State {
        Normal,
        SingleQuoted,
        DoubleQuoted,
        LineComment,
        BlockComment,
    }

    let bytes = source.as_bytes();
    // Accumulate raw bytes so multibyte UTF-8 sequences survive verbatim; the
    // scanner only ever branches on ASCII delimiters, so byte-wise copying is safe.
    let mut out: Vec<u8> = Vec::with_capacity(source.len() + 32);
    let mut index = 0usize;
    let mut state = State::Normal;

    while index < bytes.len() {
        match state {
            State::Normal => {
                if starts_with_magic_constant(bytes, index, b"__FILE__")
                    && is_identifier_boundary(bytes, index, b"__FILE__".len())
                {
                    out.extend_from_slice(file_literal.as_bytes());
                    index += b"__FILE__".len();
                    continue;
                }
                if starts_with_magic_constant(bytes, index, b"__DIR__")
                    && is_identifier_boundary(bytes, index, b"__DIR__".len())
                {
                    out.extend_from_slice(dir_literal.as_bytes());
                    index += b"__DIR__".len();
                    continue;
                }

                if bytes[index] == b'\'' {
                    state = State::SingleQuoted;
                } else if bytes[index] == b'"' {
                    state = State::DoubleQuoted;
                } else if bytes[index] == b'/'
                    && index + 1 < bytes.len()
                    && bytes[index + 1] == b'/'
                {
                    state = State::LineComment;
                } else if bytes[index] == b'#' {
                    state = State::LineComment;
                } else if bytes[index] == b'/'
                    && index + 1 < bytes.len()
                    && bytes[index + 1] == b'*'
                {
                    state = State::BlockComment;
                }

                out.push(bytes[index]);
                index += 1;
            }
            State::SingleQuoted => {
                out.push(bytes[index]);
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    out.push(bytes[index + 1]);
                    index += 2;
                    continue;
                }
                if bytes[index] == b'\'' {
                    state = State::Normal;
                }
                index += 1;
            }
            State::DoubleQuoted => {
                out.push(bytes[index]);
                if bytes[index] == b'\\' && index + 1 < bytes.len() {
                    out.push(bytes[index + 1]);
                    index += 2;
                    continue;
                }
                if bytes[index] == b'"' {
                    state = State::Normal;
                }
                index += 1;
            }
            State::LineComment => {
                out.push(bytes[index]);
                if bytes[index] == b'\n' {
                    state = State::Normal;
                }
                index += 1;
            }
            State::BlockComment => {
                out.push(bytes[index]);
                if bytes[index] == b'*' && index + 1 < bytes.len() && bytes[index + 1] == b'/' {
                    out.push(b'/');
                    index += 2;
                    state = State::Normal;
                    continue;
                }
                index += 1;
            }
        }
    }

    String::from_utf8(out).unwrap_or_else(|_| source.to_string())
}

fn starts_with_magic_constant(bytes: &[u8], index: usize, needle: &[u8]) -> bool {
    bytes.get(index..index + needle.len()) == Some(needle)
}

fn is_identifier_boundary(bytes: &[u8], index: usize, len: usize) -> bool {
    let before_ok = index == 0 || !is_php_identifier_byte(bytes[index - 1]);
    let after = index + len;
    let after_ok = after >= bytes.len() || !is_php_identifier_byte(bytes[after]);
    before_ok && after_ok
}

fn is_php_identifier_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_'
}

fn php_single_quoted_literal(value: &str) -> String {
    let escaped = value.replace('\\', "\\\\").replace('\'', "\\'");
    format!("'{}'", escaped)
}

fn absolutize_path(path: &Path) -> PathBuf {
    if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()
            .unwrap_or_else(|_| PathBuf::from("."))
            .join(path)
    }
}

pub fn into_dynamic_compilation(compiled: CompiledBundle) -> DynamicCompilation {
    DynamicCompilation {
        compile_dependencies: None,
        source_declarations: Arc::from([]),
        required_libraries: compiled.required_libraries,
        chunks: compiled.chunks,
        host_imports: compiled.host_imports,
        component_source: compiled.component_source,
        component_binary: compiled.component_binary,
        entry_path: None,
        app_shell: compiled.app_shell,
        prepared_sources: None,
    }
}

impl PhpIncludeRuntime {
    fn select_completion_permission(&mut self, language:Language) {
        if !self.completion_permission_override {
            self.completion_permission = registered_completion_lowering(language)
                .then(|| (FrontendIdentity::of(language), registered_completion_lowering as CompletionLoweringHook));
        }
        let permitted = self.completion_permission.is_some_and(|(identity, hook)|
            identity == FrontendIdentity::of(language) && hook(language));
        self.preparation_plan.set_retirement_admission(permitted);
    }
    fn trace_include_event(
        &self,
        event: &str,
        kind: &str,
        target: &str,
        resolved_path: &Path,
        canonical_path: &Path,
    ) {
        if std::env::var_os("VYBE_TRACE_PHP_INCLUDE").is_some() {
            eprintln!(
                "[php-include] event={event} depth={} kind={kind} target={} resolved={} canonical={} state={:?}",
                self.current_paths.len(),
                target,
                resolved_path.display(),
                canonical_path.display(),
                self.once_resources.get(canonical_path)
            );
        }
    }

    fn trace_include_cache_event(&self, event: &str, path: &Path, fingerprint: u64) {
        if std::env::var_os("VYBE_TRACE_PHP_INCLUDE_CACHE").is_some() {
            eprintln!(
                "[php-include-cache] event={event} fingerprint={fingerprint} path={}",
                path.display()
            );
        }
    }

    fn new(caps: Capabilities) -> Self {
        Self {
            caps,
            vm: std::ptr::null_mut(),
            current_paths: Vec::new(),
            once_resources: HashMap::new(),
            include_cache: None,
            include_scope: None,
            resident_units: HashMap::new(),
            #[cfg(feature="service-verification")] resident_hits:0,
            preparing: HashMap::new(),
            prepared_ready: HashMap::new(),
            lowering: HashMap::new(),
            retired_compiled: Default::default(),
            link_budget: Arc::new(AtomicUsize::new(0)),
            contextual_compile_failure: false,
            fused_context:None,completion_permission:None,completion_permission_override:false,
            preparation_plan: Default::default(),
            preparation_hints: HashMap::new(),
            helper_session: Default::default(),
            #[cfg(any(test,v6_service_gate))]
            completion_handoff_hits: 0,
            #[cfg(any(test,v6_service_gate))]
            current_source_rank_receipts: 0,
            #[cfg(v6_service_gate)]
            committed_foreground_receipts: 0,
            #[cfg(v6_service_gate)]
            committed_foreground_ready_paths: 0,
            builtin_callables: HashMap::new(),
            php_trait_declarations: HashMap::new(),
        }
    }

    fn begin_once_resource(&mut self, path: &Path) -> bool {
        match self.once_resources.get(path) {
            Some(IncludeOnceState::Loading | IncludeOnceState::Loaded) => false,
            None => {
                self.once_resources
                    .insert(path.to_path_buf(), IncludeOnceState::Loading);
                true
            }
        }
    }

    fn finish_once_resource(&mut self, path: PathBuf) {
        self.once_resources.insert(path, IncludeOnceState::Loaded);
    }

    fn fail_once_resource(&mut self, path: &Path) {
        self.once_resources.remove(path);
    }

    fn publish_include_resource(&mut self, path: &Path) -> bool {
        if self.once_resources.contains_key(path) {
            false
        } else {
            self.once_resources
                .insert(path.to_path_buf(), IncludeOnceState::Loading);
            true
        }
    }

    fn is_currently_loading(&self, path: &Path) -> bool {
        self.current_paths.iter().any(|active| active == path)
    }

    fn activate(
        &mut self,
        vm: &mut VM,
        entry_path: Option<&Path>,
    ) -> ActivePhpRuntimeGuard {
        self.vm = vm as *mut VM;
        self.current_paths.clear();
        self.builtin_callables.clear();
        self.resident_units.clear();
        #[cfg(feature="service-verification")] {self.resident_hits=0;}
        self.php_trait_declarations.clear();
        self.preparation_hints.clear();
        self.prepared_ready.clear();
        self.lowering.clear();self.retired_compiled.clear();
        if let Some(path) = entry_path {
            let canonical_path = path.canonicalize().unwrap_or_else(|_| path.to_path_buf());
            self.current_paths.push(canonical_path.clone());
            self.once_resources
                .insert(canonical_path, IncludeOnceState::Loaded);
        }
        let previous = ACTIVE_PHP_RUNTIME.with(|slot| slot.replace(Some(self as *mut _)));
        ActivePhpRuntimeGuard { previous }
    }

    /// PHP does NOT report every include failure the same way, so neither do
    /// we: a file that cannot be opened is a warning and `false` for
    /// `include`/`include_once` but FATAL for `require`/`require_once`, and an
    /// error raised *inside* the included file is the script's error either
    /// way — never "the include returned false". Every `Err` returned here is
    /// thrown by the host-fn wrapper in `ensure_php_runtime_registered`.
    fn handle_dynamic_include(&mut self, args: &[Value]) -> Result<Value, String> {
        let kind = args.first().map(value_to_string).unwrap_or_default();
        let fatal_if_missing = matches!(kind.as_str(), "require" | "require_once");

        if !self.caps.has(Capability::DynamicCompile) {
            return Err(format!(
                "{kind} needs the DynamicCompile capability, which this run was not granted"
            ));
        }

        if self.vm.is_null() {
            return Err(format!("{kind} ran with no active VM"));
        }

        let vm = unsafe { &mut *self.vm };

        let target = args.get(1).map(value_to_string).unwrap_or_default();
        let caller = self
            .current_paths
            .last()
            .cloned()
            .ok_or_else(|| "dynamic PHP include has no active caller path".to_string())?;
        let entry = self
            .current_paths
            .first()
            .cloned()
            .ok_or_else(|| "dynamic PHP include has no active entry path".to_string())?;
        let resolved_path = resolve_php_include_path(&entry, &caller, &target);
        let canonical_path = resolved_path
            .canonicalize()
            .unwrap_or_else(|_| resolved_path.clone());
        let _include_phase = DebugPhase::new(
            vm.debug_report(),
            format!("include {}", canonical_path.display()),
        );

        let load_once = matches!(kind.as_str(), "include_once" | "require_once");
        if load_once && !self.begin_once_resource(&canonical_path) {
            self.trace_include_event("skip-once", &kind, &target, &resolved_path, &canonical_path);
            return Ok(Value::I32(1));
        }
        if !load_once && self.is_currently_loading(&canonical_path) {
            self.trace_include_event(
                "skip-active",
                &kind,
                &target,
                &resolved_path,
                &canonical_path,
            );
            return Ok(Value::I32(1));
        }
        let published_plain_resource = if load_once {
            false
        } else {
            self.publish_include_resource(&canonical_path)
        };
        self.trace_include_event("load", &kind, &target, &resolved_path, &canonical_path);

        let language = languages::find_by_name("php")
            .ok_or_else(|| "php language profile missing".to_string())?;
        self.refresh_php_preparation_hints(&caller, &entry, language, vm);

        // Reuse an earlier compilation of this exact file only where everything
        // the compile READ is still what it was — the file itself and whatever
        // it statically pulls in (the dependency set), plus the entry path
        // (the fingerprint). Imports are relocated against the live VM.
        let cache = self.include_cache.clone();
        let fingerprint = cache
            .as_ref()
            .map(|_| module_fingerprint(vm, &entry))
            .unwrap_or(0);
        let request_resident=self.resident_requested(language);
        if let Some(resident)=self.resident_hit(vm,&canonical_path,language,&entry,&canonical_path) {
            let imports=vm.request_unit_imports(&resident.unit).ok().and_then(|declared|resolve_imports(vm,declared).ok());
            if let Some(imports)=imports.filter(|imports|vm.request_unit_accepts_imports(&resident.unit,imports)) {
                #[cfg(feature="service-verification")] {self.resident_hits+=1;}
                self.trace_include_cache_event("resident-hit",&canonical_path,fingerprint);
                self.preparing.remove(&canonical_path);self.prepared_ready.remove(&canonical_path);
                self.lowering.remove(&canonical_path);self.replenish_php_preparation();
                return self.run_resident_include(vm,resident,canonical_path,
                    load_once || published_plain_resource,args.get(2).cloned(),imports);
            }
            // Denial is still before code publication, initializer or guest execution.
            self.resident_units.remove(&canonical_path);
        }
        if let Some(cache) = cache.as_ref() {
            if let Some(hit) = cache.get(&canonical_path, fingerprint).filter(|hit| {
                match hit { Ok(compiled) => compiled.libraries_available(vm), Err(_) => true }
            }) {
                self.preparing.remove(&canonical_path);
                self.prepared_ready.remove(&canonical_path);
                self.lowering.remove(&canonical_path);
                self.replenish_php_preparation();
                self.trace_include_cache_event("hit", &canonical_path, fingerprint);
                // A cached compile FAILURE is still that file's answer; letting
                // it fall through would recompile a broken include on every
                // request, which is exactly what the entry cache refuses to do.
                let compiled = match hit {
                    Ok(compiled) => compiled,
                    Err(err) => {
                        if load_once || published_plain_resource {
                            self.fail_once_resource(&canonical_path);
                        }
                        return Err(err);
                    }
                };
                return self.run_included_compilation(
                    vm,
                    compiled,
                    kind,
                    canonical_path,
                    load_once || published_plain_resource,
                    args.get(2).cloned(),
                );
            }
        }

        // A cache hit has already validated the source dependency set. Only
        // open and allocate the source text when a fresh compilation is needed.
        let source_before=request_resident.then(||crate::bundle::SourceObservation::capture(&resolved_path));
        let source = match fs::read_to_string(&resolved_path) {
            Ok(source) => source,
            Err(e) => {
                if load_once || published_plain_resource {
                    self.fail_once_resource(&canonical_path);
                }
                let detail = format!(
                    "{kind}({}): failed to open stream: {e}",
                    resolved_path.display()
                );
                if fatal_if_missing {
                    return Err(detail);
                }
                eprintln!("Warning: {detail}");
                return Ok(Value::Bool(false));
            }
        };

        let bundle = bundle_from_source(source, language, resolved_path.clone());
        let mut compile=|| {
            if cache.is_some() || request_resident {
                crate::bundle::record_source_reads(||self.compile_pending_with_reuse(vm,&bundle,&entry,true))
            } else {(self.compile_pending_with_reuse(vm,&bundle,&entry,false),None)}
        };
        let ((result,deps),observations)=if request_resident {
            crate::bundle::record_source_observations(compile)
        } else {(compile(),None)};
        let deps=deps.map(|mut paths|{paths.push(resolved_path.clone());paths});
        let deps=if result.is_err() && self.contextual_compile_failure {None}else{deps};
        // Existing raw cache storage is retained on all capture denials and errors.
        if let Some(cache)=&cache {
            let cache_span=crate::compilation_timeline::span("include_cache_store",&canonical_path);
            cache.store_with_chunks(&canonical_path,fingerprint,deps,
                result.as_ref().map(|pending|(&pending.compilation,pending.code.chunks())).map_err(|err|err.as_str()));
            drop(cache_span);self.trace_include_cache_event("store",&canonical_path,fingerprint);
        }
        let compiled=match result {Ok(compiled)=>compiled,Err(err)=>{
            if load_once || published_plain_resource {self.fail_once_resource(&canonical_path);}return Err(err);
        }};
        let resident_source=if request_resident {SourceReceipt::capture(source_before,observations)}else{None};
        self.run_pending_included_compilation_recorded(vm,compiled,kind,canonical_path,
            load_once || published_plain_resource,args.get(2).cloned(),
            resident_source.map(|source|(source,language,entry,cache.clone().map(|cache|(cache,fingerprint)))))
    }

    /// Install a compiled include into the live VM and run it — the half that
    /// is identical whether the compilation was just produced or came from the
    /// cache.
    fn run_included_compilation(
        &mut self,
        vm: &mut VM,
        compiled: DynamicCompilation,
        _kind: String,
        canonical_path: PathBuf,
        publishes_once_resource: bool,
        explicit_scope: Option<Value>,
    ) -> Result<Value, String> {
        self.run_pending_included_compilation(vm, PendingCompilation::raw(compiled), _kind,
            canonical_path, publishes_once_resource, explicit_scope)
    }

    fn run_pending_included_compilation(
        &mut self, vm: &mut VM, pending: PendingCompilation, _kind: String,
        canonical_path: PathBuf, publishes_once_resource: bool, explicit_scope: Option<Value>,
    ) -> Result<Value, String> {
        self.run_pending_included_compilation_recorded(vm,pending,_kind,canonical_path,
            publishes_once_resource,explicit_scope,None)
    }

    fn run_pending_included_compilation_recorded(
        &mut self, vm: &mut VM, pending: PendingCompilation, _kind: String,
        canonical_path: PathBuf, publishes_once_resource: bool, explicit_scope: Option<Value>,
        resident_source:Option<(SourceReceipt,Language,PathBuf,Option<(Arc<dyn IncludeCompileCache>,u64)>)>,
    ) -> Result<Value, String> {
        let PendingCompilation { compilation: compiled, code } = pending;
        // A function-scoped include inherits the caller's symbol table. Keep
        // it distinct from request globals: functions declared/called by the
        // template must still see genuine `global $name` bindings.
        let caller_locals = if explicit_scope.is_some() { Vec::new() } else {
            vm.current_source_scope()
                .filter(|(name, _)| name != "<script>")
                .map(|(_, locals)| locals).unwrap_or_default()
        };
        let base_chunk_index = vm.chunks.len();
        let resident_capture=resident_source.map(|(source,language,entry,raw_cache)|{
            let callables=code.chunks().iter().enumerate().filter(|(_,c)|should_publish_chunk_name(&c.name))
                .map(|(i,c)|request_code_space::Callable{name:c.name.clone(),arity:c.arity,chunk:base_chunk_index+i}).collect();
            (source,FrontendIdentity::of(language),entry,callables,raw_cache)
        });
        let install_span = crate::compilation_timeline::span("include_install", &canonical_path);
        crate::host_imports::install(vm, &compiled.host_imports);
        install_chunk_globals(vm, code.chunks(), base_chunk_index);

        let child_active_imports = code.chunks()
            .first()
            .map(|chunk| chunk.imports.as_slice())
            .unwrap_or(&[]);
        let child_active_resolved_imports = resolve_imports(vm, child_active_imports)?;
        drop(install_span);

        #[cfg(feature="service-verification")]if resident_capture.is_some(){eprintln!("resident attempt: {:?}",vybe_runtime::vm::LoadedRequestUnit::admission(code.chunks()));}else{eprintln!("resident no source receipt/capability");}
        let mut loaded=None;
        let result=self.execute_include_scope(vm,&compiled.source_declarations,canonical_path.clone(),
            publishes_once_resource,explicit_scope,caller_locals,|vm| {
                match code {
                    LinkPayload::Raw(chunks)=>if resident_capture.is_some(){
                        let(result,unit)=vm.run_linked_nested_request(chunks,child_active_resolved_imports);loaded=unit;result
                    }else{vm.run_linked_nested(chunks,child_active_resolved_imports)},
                    LinkPayload::Prepared(code,permit)=>{
                        drop(permit);
                        if resident_capture.is_some(){let(result,unit)=vm.run_linked_nested_prepared_request(code,child_active_resolved_imports);loaded=unit;result}
                        else{vm.run_linked_nested_prepared(code,child_active_resolved_imports)}
                    }
                }
            });
        #[cfg(feature="service-verification")]eprintln!("resident loader handle={}",loaded.is_some());
        if let(Some(unit),Some((source,frontend,entry,callables,raw_cache)))=(loaded,resident_capture){
            // Execution only borrowed the compilation's metadata. Transfer the
            // original owner after that borrow ends instead of cloning it for
            // every cold include before resident admission is known.
            if let Some(resident)=ResidentEntry::capture(vm,unit,source,frontend,entry,compiled,callables){
                // Keep the cold raw snapshot until execution AND all resident
                // receipts succeed. Errors never retry and preserve fallback.
                let retire=result.is_ok();
                self.resident_units.insert(canonical_path.clone(),Arc::new(resident));
                if retire {
                    if let Some((cache,fingerprint))=raw_cache.filter(|(cache,_)|cache.request_local_only()) {
                        cache.retire_request_snapshot(&canonical_path,fingerprint);
                    }
                }
            }
        }
        result
    }

    fn execute_include_scope(&mut self,vm:&mut VM,source_declarations:&[Statement],canonical_path:PathBuf,
        publishes_once_resource:bool,explicit_scope:Option<Value>,caller_locals:Vec<(u16,String,Value)>,
        execute:impl FnOnce(&mut VM)->Result<Value,vybe_runtime::VMError>)->Result<Value,String> {
        let scope_span = crate::compilation_timeline::span("include_scope_materialization", &canonical_path);
        let previous_scope = self.include_scope.clone();
        if let Some(scope) = explicit_scope {
            self.include_scope = Some(scope);
        } else if !caller_locals.is_empty() {
            let mut scope = Object::new();
            for (_, name, value) in &caller_locals {
                scope.properties.insert(name.trim_start_matches('$').to_string(), value.clone());
            }
            self.include_scope = Some(Value::Object(vybe_runtime::heap::alloc(scope)));
        }
        drop(scope_span);
        let declarations_span = crate::compilation_timeline::span("include_declarations", &canonical_path);
        self.install_source_declarations(source_declarations);
        drop(declarations_span);
        self.current_paths.push(canonical_path.clone());
        let result=execute(vm).map_err(|err|err.to_string());
        self.current_paths.pop();
        if !caller_locals.is_empty() {
            if let Some(Value::Object(scope)) = &self.include_scope {
                let scope = scope.lock().unwrap();
                for (slot, name, _) in &caller_locals {
                    if let Some(value) = scope.properties.get(name.trim_start_matches('$')) {
                        vm.set_current_source_local(*slot, value.clone());
                    }
                }
            }
        }
        self.include_scope = previous_scope;

        match result {
            Ok(Value::Null) => {
                if publishes_once_resource {
                    self.finish_once_resource(canonical_path);
                }
                Ok(Value::I32(1))
            }
            Ok(value) => {
                if publishes_once_resource {
                    self.finish_once_resource(canonical_path);
                }
                Ok(value)
            }
            // The included file ran and failed. That is the script's error, not
            // a failed file open — reporting it as `false` made a fatal inside
            // an include indistinguishable from a missing file.
            Err(err) => {
                if publishes_once_resource {
                    self.fail_once_resource(&canonical_path);
                }
                Err(err)
            }
        }
    }

    /// Materialize a compiler-lowered builtin only when a dynamic callback
    /// actually requests it. Reuse the normal PHP lowering and cache the
    /// resulting function in this request's live VM, indexed by argument count.
    fn builtin_callable(&mut self, name: &str, argc: usize) -> Result<Value, String> {
        let name = name.trim_start_matches('\\').to_ascii_lowercase();
        let key = (name.clone(), argc);
        if let Some(value) = self.builtin_callables.get(&key) {
            return Ok(value.clone());
        }
        if argc > 255 || name.is_empty()
            || !name.bytes().all(|b| b.is_ascii_alphanumeric() || b == b'_') {
            return Ok(Value::Null);
        }
        let language = languages::find_by_name("php").ok_or("PHP language unavailable")?;
        let profile = crate::profile::parse_profile_shared((language.profile_source)())?;
        if profile.lookup_builtin(&name).is_none() {
            return Ok(Value::Null);
        }
        let parameters = (0..argc).map(|i| format!("$__arg{i}")).collect::<Vec<_>>().join(", ");
        let source = format!("<?php return function({parameters}) {{ return \\{name}({parameters}); }};");
        let path = PathBuf::from(format!("<php-builtin-{name}-{argc}>.php"));
        let bundle = bundle_from_source(source, language, &path);
        let vm = unsafe { &mut *self.vm };
        let compiled = self.compile_dynamic_php(vm, &bundle, &path)?;
        let value = self.run_included_compilation(vm, compiled, "builtin".into(), path, false, None)?;
        self.builtin_callables.insert(key, value.clone());
        Ok(value)
    }

    fn install_source_declarations(&mut self, declarations: &[Statement]) {
        for statement in declarations {
            if let StmtKind::ClassDecl { name, modifiers, .. } = &statement.kind {
                if modifiers.kind == ClassKind::Trait {
                    self.php_trait_declarations.insert(name.to_ascii_lowercase(), statement.clone());
                }
            }
        }
    }

    fn include_available_php_traits(
        &mut self,
        module: &mut Module
    ) {
        fn class_statements(body: &[Statement]) -> Vec<&Statement> {
            let mut classes = Vec::new();
            for statement in body {
                match &statement.kind {
                    StmtKind::ClassDecl { .. } => classes.push(statement),
                    StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
                        classes.extend(class_statements(body));
                    }
                    _ => {}
                }
            }
            classes
        }
        fn trait_name(statement: &Statement) -> Option<&str> {
            match &statement.kind {
                StmtKind::ClassDecl { name, modifiers, .. }
                    if modifiers.kind == ClassKind::Trait => Some(name),
                _ => None,
            }
        }
        fn trait_uses(statement: &Statement) -> Vec<(String, String)> {
            match &statement.kind {
                StmtKind::ClassDecl { name, members, .. } => members
                    .iter()
                    .filter_map(|member| match member {
                        ClassMember::Augment(augment) => Some((name.clone(), augment.from.clone())),
                        _ => None,
                    })
                    .collect(),
                _ => Vec::new(),
            }
        }
        let classes = class_statements(&module.body);
        let mut pending: Vec<(String, String)> = classes.iter().flat_map(|statement| trait_uses(statement)).collect();
        if pending.is_empty() {return;}
        // The original module remains unchanged until the final splice. Do not
        // add imported schemas here: their existing ordered resolution uses
        // the separate pending/seen lists.
        let local_traits: HashSet<String> = classes.iter()
            .filter_map(|statement| trait_name(statement))
            .map(str::to_ascii_lowercase).collect();
        let mut inserted = Vec::new();
        let mut seen = HashSet::new();
        let mut resolved_augments = HashMap::new();
        while let Some((_owner, source_name)) = pending.pop() {
            let source_key = source_name.to_ascii_lowercase();
            if !seen.insert(source_key.clone()) {
                continue;
            }
            if local_traits.contains(&source_key) {
                continue;
            }
            let known = self.php_trait_declarations.get(&source_key).cloned().or_else(|| {
                let suffix = format!(".{source_key}");
                let mut matches = self.php_trait_declarations.iter()
                    .filter(|(name, _)| name.ends_with(&suffix));
                match (matches.next(), matches.next()) {
                    (Some((_, statement)), None) => Some(statement.clone()),
                    _ => None,
                }
            });
            let declaration = known;
            if let Some(declaration) = declaration {
                if let Some(name) = trait_name(&declaration) {
                    resolved_augments.insert(source_key, name.to_string());
                }
                pending.extend(trait_uses(&declaration));
                inserted.push(declaration);
            }
        }
        fn rewrite_augments(body: &mut [Statement], resolved: &HashMap<String, String>) {
            for statement in body {
                match &mut statement.kind {
                    StmtKind::ClassDecl { members, .. } => {
                        for member in members {
                            if let ClassMember::Augment(augment) = member {
                                if let Some(name) = resolved.get(&augment.from.to_ascii_lowercase()) {
                                    augment.from = name.clone();
                                }
                            }
                        }
                    }
                    StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
                        rewrite_augments(body, resolved)
                    }
                    _ => {}
                }
            }
        }
        rewrite_augments(&mut module.body, &resolved_augments);
        rewrite_augments(&mut inserted, &resolved_augments);
        if !inserted.is_empty() {
            inserted.reverse();
            module.body.splice(0..0, inserted);
        }
    }

    fn cancel_request_preparation(&mut self){
        self.preparing.clear();self.lowering.clear();self.retired_compiled.clear();self.prepared_ready.clear();self.preparation_hints.clear();
        self.preparation_plan=Default::default();
    }
    fn prepare_php_source_hints(&mut self, bundle: &Bundle, entry: &Path, vm: Option<&VM>) {
        if crate::frontend_preparation::for_parser(bundle.language.parse).is_some_and(|hook|!hook.early_hints){return;}
        let fused = self.fused_context.is_some();
        if !self.php_preparation_enabled() { return; }
        let Some(source) = bundle.sources.first() else { return; };
        let caller = absolutize_path(&source.path);
        let key = caller.canonicalize().unwrap_or_else(|_| caller.clone());
        // A new compilation can follow an edit or a repeated plain include.
        // Replace its old hints rather than retaining a stale source plan.
        let mut hints = if fused {
            crate::php_include_hints::IncludeHints::immediate(&source.code, &caller)
        } else {
            crate::php_include_hints::IncludeHints::new(&source.code, &caller)
        };
        let paths = hints.take_ready_paths(|name| {
            match vm.and_then(|vm| vm.global(name)) {
                Some(Value::String(value)) => Some(value.to_string()),
                _ => None,
            }
        })
            .into_iter().map(|path| resolve_prepared_php_include_path(entry, &source.path, &path, fused));
        self.preparation_hints.remove(&key);
        if !crate::preparation::DeferredPreparationHints::is_empty(&hints) {
            self.preparation_hints.insert(key, Box::new(hints));
        }
        self.queue_current_source_preparation(paths, bundle.language,&caller,true);
    }

    fn refresh_php_preparation_hints(&mut self, caller: &Path, entry: &Path,
        language: Language, vm: &VM,
    ) {
        let fused = self.fused_context.is_some();
        if !self.php_preparation_enabled() { return; }
        let Some(hints) = self.preparation_hints.get_mut(caller) else {
            if fused {
                // Fully drained parents still own retained sibling identities.
                // Selecting that owner restores its row without another scout,
                // VM query, pending-queue mutation or default-path refill.
                let owner=caller.canonicalize().unwrap_or_else(|_|absolutize_path(caller));
                self.preparation_plan.prioritize_current(owner,&[],false);
            }
            return;
        };
        let paths = hints.take_ready_paths(&mut |name| match vm.global(name) {
            Some(Value::String(value)) => Some(value.to_string()),
            _ => None,
        }).into_iter().map(|path| resolve_prepared_php_include_path(entry, caller, &path, fused));
        if hints.is_empty() { self.preparation_hints.remove(caller); }
        self.queue_current_source_preparation(paths, language,caller,false);
    }

    fn refresh_worker_preparation_hints(&mut self, entry: &Path, language: Language, vm: &VM) {
        if !self.php_preparation_enabled() { return; }
        let discoveries: Vec<_> = self.preparing.values().flat_map(|ticket|std::iter::from_fn(||ticket.take_hints())).collect();
        self.consume_worker_preparation_hints(discoveries, entry, language, vm);
    }

    fn consume_worker_preparation_hints(&mut self,
        discoveries:impl IntoIterator<Item=(PathBuf,Box<dyn crate::preparation::DeferredPreparationHints>)>,
        entry:&Path,language:Language,vm:&VM,
    ) {
        let paths=self.resolve_preparation_hint_rows(discoveries,entry,language,vm);
        self.queue_php_preparation(paths,language);
    }

    // This receipt belongs to the foreground source being parsed, not an
    // arbitrary speculative worker. Resolve once through the same VM queries.
    fn consume_current_source_preparation_hints(&mut self,caller:PathBuf,
        hints:Box<dyn crate::preparation::DeferredPreparationHints>,
        entry:&Path,language:Language,vm:&VM,
    ) {
        #[cfg(v6_service_gate)] { self.committed_foreground_receipts += 1; }
        let paths=self.resolve_preparation_hint_rows(std::iter::once((caller.clone(),hints)),entry,language,vm);
        #[cfg(v6_service_gate)] { self.committed_foreground_ready_paths += paths.len(); }
        self.queue_current_source_preparation(paths,language,&caller,true);
    }

    fn resolve_preparation_hint_rows(&mut self,
        discoveries: impl IntoIterator<Item = (PathBuf, Box<dyn crate::preparation::DeferredPreparationHints>)>,
        entry: &Path, language: Language, vm: &VM,
    ) -> Vec<PathBuf> {
        let fused = self.fused_context.is_some();
        let capabilities=crate::frontend_preparation::snapshot(language.parse).capabilities;
        let mut ready = Vec::new();
        for (caller, mut hints) in discoveries {
            crate::compilation_timeline::event("source_hints", "consumed", &caller);
            ready.extend(hints.take_ready_paths(&mut |name| match vm.global(name) {
                Some(Value::String(value)) => Some(value.to_string()),
                _ => None,
            }).into_iter().map(|path| {
                if let Some(resolve)=capabilities.resolve_path {resolve(entry,&caller,&path,fused)}
                else if capabilities.prepare.is_some(){
                    let path=PathBuf::from(path);if path.is_absolute(){path}else{caller.parent().unwrap_or(Path::new(".")).join(path)}
                }else{resolve_prepared_php_include_path(entry,&caller,&path,fused)}
            }));
            self.preparation_hints.remove(&caller);
            if !hints.is_empty() { self.preparation_hints.insert(caller, hints); }
        }
        ready
    }

    fn php_preparation_enabled(&self) -> bool {
        self.caps.has(Capability::DynamicCompile)
            && !std::env::var_os("VYBE_PARALLEL_PREPARATION").is_some_and(|value| value == "0")
    }

    fn queue_php_preparation(&mut self, paths: impl IntoIterator<Item = PathBuf>, language: Language) {
        for path in paths {
            let Ok(key) = path.canonicalize() else { continue; };
            if self.preparing.contains_key(&key) || self.prepared_ready.contains_key(&key) || self.lowering.contains_key(&key) || self.once_resources.contains_key(&key) { continue; }
            let parse_path=if self.fused_context.is_some(){path}else{key.clone()};
            self.preparation_plan.enqueue_as(key, parse_path, language);
        }
        self.replenish_php_preparation();
    }

    fn queue_current_source_preparation(&mut self,
        paths: impl IntoIterator<Item = PathBuf>, language: Language,
        caller:&Path, full:bool,
    ) {
        if self.fused_context.is_none() {
            self.queue_php_preparation(paths, language);
            return;
        }
        let mut priority=Vec::new();
        let paths: Vec<_> = paths.into_iter().filter_map(|path| {
            let key = path.canonicalize().ok()?;
            if self.once_resources.contains_key(&key) { return None; }
            if priority.len()<128&&!priority.contains(&key) { priority.push(key.clone()); }
            if self.preparing.contains_key(&key)||self.prepared_ready.contains_key(&key)
                ||self.lowering.contains_key(&key) { return None; }
            Some((key,path,language))
        }).collect();
        let owner=caller.canonicalize().unwrap_or_else(|_|absolutize_path(caller));
        #[cfg(any(test,v6_service_gate))] { if self.preparation_plan.retirement_rank_gate_enabled() { self.current_source_rank_receipts += 1; } }
        self.preparation_plan.prioritize_current(owner,&priority,full);
        self.preparation_plan.prioritize_as(paths);
        self.replenish_php_preparation();
    }

    fn prioritize_resolved_source_hints(&mut self,
        hints: Option<&crate::preparation::ResolvedPreparationHints>, caller: &Path,
        entry: &Path, language: Language,
    ) {
        let Some(paths) = hints.and_then(|hints| hints.lock().ok().map(|paths| paths.clone())) else { return; };
        let fused = self.fused_context.is_some();
        self.queue_current_source_preparation(paths.into_iter().map(|path|
            resolve_prepared_php_include_path(entry, caller, &path, fused)), language,caller,true);
    }

    fn replenish_php_preparation(&mut self) {
        if !self.php_preparation_enabled() { return; }
        let Some(pool) = crate::preparation::PreparationPool::shared() else { return; };
        let loaded = &self.once_resources;
        if let Some(context)=self.fused_context.clone(){
            // A fused job owns one preparation slot, not two slots merely
            // because its lowering-result receiver has already been installed.
            let staged_only=self.lowering.keys().filter(|p|!self.preparing.contains_key(*p)).count();
            let window=4usize.saturating_sub(self.prepared_ready.len()+staged_only);
            let lowerings=&mut self.lowering;
            let link_budget=self.link_budget.clone();
            self.preparation_plan.replenish_by(&mut self.preparing,window,|key,path,language|{
                let (send,result)=crossbeam_channel::bounded(1);
                let cancelled=Arc::new(std::sync::atomic::AtomicBool::new(false));
                let work_cancelled=cancelled.clone();let context=context.clone();let link_budget=link_budget.clone();
                let ticket_budget=link_budget.clone();
                let post:crate::preparation::PostPrepare=Box::new(move|prepared|{
                    if work_cancelled.load(std::sync::atomic::Ordering::Relaxed){return;}
                    let lowered=context.lower_prepared(prepared,language,&link_budget);
                    if !work_cancelled.load(std::sync::atomic::Ordering::Relaxed){if let Some(lowered)=lowered{let _=send.send(lowered);}}
                });
                let ticket=pool.submit_with_post_prepare(path.clone(),language,Some(crate::php_include_hints::discover_immediate),Some(post))?;
                lowerings.insert(key,LoweringTicket{result,cancelled,link_budget:ticket_budget});Some(ticket)
            },|path|loaded.contains_key(path));
        }else{
            self.preparation_plan.replenish_with_hints(pool, &mut self.preparing, 4usize.saturating_sub(self.prepared_ready.len()+self.lowering.len()),
                Some(crate::php_include_hints::discover),|path| loaded.contains_key(path));
        }
    }

    fn prepare_known_php_includes(&mut self, module: &Module, bundle: &Bundle, entry: &Path) {
        if !self.caps.has(Capability::DynamicCompile)
            || std::env::var_os("VYBE_PARALLEL_PREPARATION").is_some_and(|value| value == "0")
            || !bundle.sources.iter().any(|source|
                source.code.contains("include") || source.code.contains("require"))
        {
            return;
        }
        let Some(caller) = bundle.sources.first().map(|source| source.path.as_path()) else { return; };
        fn literal_path(expr: &vybe_ast::Expression) -> Option<String> {
            use vybe_ast::{ExprKind, Literal, BinOp};
            match &expr.kind {
                ExprKind::Lit(Literal::Str(path)) => Some(path.clone()),
                ExprKind::Binary { op: BinOp::Concat, left, right } =>
                    Some(literal_path(left)? + &literal_path(right)?),
                _ => None,
            }
        }
        let mut paths = Vec::new();
        for statement in &module.body {
            walk_include_lookahead_exprs(statement, self.fused_context.is_none(), &mut |expr| {
                use vybe_ast::ExprKind;
                if let ExprKind::Call { callee, args, .. } = &expr.kind {
                    if matches!(&callee.kind, ExprKind::Ident(name) if name == "__php_dynamic_include") {
                        if let Some(path) = args.get(1).and_then(|arg| literal_path(&arg.value)) {
                            paths.push(resolve_php_include_path(entry, caller, &path));
                        }
                    }
                }
            });
        }
        if paths.is_empty() { return; }
        self.queue_current_source_preparation(paths, bundle.language,caller,true);
    }

    fn stage_ready_lowering(&mut self,vm:&VM,language:Language,entry:&Path) {
        // Explicit experimental gate; ordinary PEG/frontends remain unchanged.
        if std::env::var_os("VYBE_PARALLEL_LOWERING").is_none_or(|v|v!="1") {return;}
        let Some(pool)=crate::preparation::PreparationPool::shared() else {return;};
        self.stage_ready_lowering_on(pool,vm,language,entry);
    }
    fn stage_ready_lowering_on(&mut self,pool:&crate::preparation::PreparationPool,vm:&VM,language:Language,entry:&Path) {
        // Trace output describes foreground order; keep its diagnostics there.
        if crate::primitives::dump_classes_enabled() {return;}
        if ["VYBE_TRACE","VYBE_DEBUG_IMPORTS","VYBE_SLOT_AUDIT"].iter().any(|key|std::env::var_os(key).is_some()) {return;}
        // Freeze the immutable capability registry on this thread before work
        // runs elsewhere. Its OnceLock contains no per-request mutable state.
        let _=crate::primitives::instructions::host::CapabilityContext::get();
        crate::ensure_languages_registered();
        vybe_runtime::registry::register_all_platform_trees();
        vybe_runtime::registry::register_all_trees();
        let ready:Vec<_>=self.preparing.iter().filter(|(path,t)|!self.lowering.contains_key(*path)&&t.is_ready()).map(|(p,_)|p.clone()).take(4usize.saturating_sub(self.lowering.len())).collect();
        let mut discoveries=Vec::new();
        let mut module_snapshot=None;
        for path in ready {
            let Some(ticket)=self.preparing.remove(&path) else {continue;};
            while let Some(hints)=ticket.take_hints() {discoveries.push(hints);}
            let Some(mut prepared)=ticket.wait_with_early_hints(|hints|discoveries.push(hints)) else {continue;};
            if !prepared.matches_frontend(language)
                || !prepared.dependencies.as_ref().is_some_and(Vec::is_empty)
                || !prepared.module.as_ref().is_ok_and(independent_lowering) {
                self.prepared_ready.insert(path,prepared);continue;
            }
            let profile_source=(language.profile_source)();
            let Ok(profile)=crate::bundle::cached_language_profile(profile_source) else {self.prepared_ready.insert(path,prepared);continue;};
            let (sender,result)=crossbeam_channel::bounded(1);
            let cancelled=Arc::new(std::sync::atomic::AtomicBool::new(false));
            let worker_cancelled=cancelled.clone();
            // All jobs admitted in this VM-thread batch consult one immutable
            // export view; do not clone the same graph for each sibling.
            let modules=module_snapshot.get_or_insert_with(|| Arc::new(vm.modules.iter().map(|(name,record)| {
                let mut copy=vybe_runtime::ModuleRecord::new_synthetic(name.clone());
                copy.exports=record.exports.clone();(name.clone(),copy)
            }).collect::<HashMap<_,_>>())).clone();
            let mut libraries=Vec::new();
            // Record predicates without rewriting linkage until admission.
            // A rejected job retains the original, already validated AST.
            crate::preparation::bind_installed_libraries(prepared.module.as_mut().unwrap(),|name,identity| {
                let installed=library_installed(vm,name,identity);
                libraries.push((name.into(),identity.into(),installed));false
            });
            let worker_frontend=prepared.frontend_identity;
            let worker_registration=crate::frontend_preparation::snapshot(language.parse);
            if worker_frontend!=FrontendIdentity::with_registration(language,worker_registration) {
                self.prepared_ready.insert(path,prepared);continue;
            }
            let helpers=self.helper_session.selected(worker_registration.helper_selection());
            let input=Arc::new(std::sync::Mutex::new(Some(prepared)));
            let worker_input=input.clone();
            let worker_path=path.clone();
            let link_budget=self.link_budget.clone();
            let ticket_budget=link_budget.clone();
            let work=Box::new(move || {
                if worker_cancelled.load(std::sync::atomic::Ordering::Relaxed) {return;}
                let prepared=worker_input.lock().unwrap().take().unwrap();
                let source=prepared.source;
                let mut module=prepared.module.unwrap();
                let mut decisions=libraries.iter();
                crate::preparation::bind_installed_libraries(&mut module,|name,identity| {
                    let (expected_name,expected_identity,installed)=decisions.next().unwrap();
                    debug_assert_eq!(name,expected_name);
                    debug_assert_eq!(identity,expected_identity);
                    *installed
                });
                debug_assert!(decisions.next().is_none());
                let started=std::time::Instant::now();
                let certificate=Arc::new(std::sync::Mutex::new(crate::compile_dependencies::Certificate::speculative()));
                let compilation=Bundle::compile_source_ast_certified_owned(language,&worker_path,module,&modules,Some(certificate.clone()),Some(profile),helpers.as_ref(),worker_registration.argument_selection()).map(|compiled|DynamicCompilation {
                    compile_dependencies: Some(Arc::new(CompilationDependencies{certificate:certificate.clone(),libraries,profile:profile_source,profile_source:language.profile_source,language_name:language.name,frontend:worker_frontend})),
                    source_declarations: Arc::from([]),required_libraries:compiled.required_libraries,
                    chunks:compiled.chunks,host_imports:compiled.host_imports,component_source:compiled.component_source,component_binary:compiled.component_binary,entry_path:Some(worker_path.clone()),
                    app_shell:compiled.app_shell,prepared_sources:None,
                });
                if !worker_cancelled.load(std::sync::atomic::Ordering::Relaxed) {
                    let compilation=compilation.map(|c| PendingCompilation::prepare(c, &link_budget));
                    let _=sender.send(LoweredSource{elapsed:started.elapsed(),path:worker_path,source,compilation});
                }
            });
            // A full queue declines only lowering, preserving preparation.
            if pool.submit_lowering(work) {
                self.lowering.insert(path,LoweringTicket{result,cancelled,link_budget:ticket_budget});
            } else if let Some(prepared)=input.lock().unwrap().take() {
                self.prepared_ready.insert(path,prepared);
            }
        }
        self.consume_worker_preparation_hints(discoveries,entry,language,vm);
    }

    fn compile_dynamic_php(
        &mut self,
        vm: &mut VM,
        bundle: &Bundle,
        entry_path: &Path,
    ) -> Result<DynamicCompilation, String> {
        self.compile_dynamic_php_with_reuse(vm, bundle, entry_path, true)
    }

    // Immediate foreground execution uses this VM's current context and never
    // publishes the output to an include cache. Speculative results still carry
    // their certificate and are validated before this function accepts them.
    fn compile_dynamic_php_with_reuse(
        &mut self,
        vm: &mut VM,
        bundle: &Bundle,
        entry_path: &Path,
        retain_dependencies: bool,
    ) -> Result<DynamicCompilation, String> {
        self.compile_pending_with_reuse(vm, bundle, entry_path, retain_dependencies)
            .map(PendingCompilation::into_raw)
    }

    fn compile_pending_with_reuse(
        &mut self, vm: &mut VM, bundle: &Bundle, entry_path: &Path,
        retain_dependencies: bool,
    ) -> Result<PendingCompilation, String> {
        let _cpu_scope = crate::compilation_timeline::span("include_compile_service",
            bundle.sources.first().map_or(entry_path, |source| source.path.as_path()));
        let registration=crate::frontend_preparation::snapshot(bundle.language.parse);
        let helpers=self.helper_session.selected(registration.helper_selection());
        let mut compiled_frontend=FrontendIdentity::with_registration(bundle.language,registration);
        self.contextual_compile_failure=false;
        // Cached entry execution can create a fresh service without calling
        // compile_bundle. Select the actual include frontend here as well.
        self.select_completion_permission(bundle.language);
        let permitted=self.completion_permission.is_some_and(|(identity,hook)|identity==FrontendIdentity::of(bundle.language)&&hook(bundle.language));
        if !permitted{self.lowering.clear();self.retired_compiled.clear();self.fused_context=None;}
        self.retired_compiled.select(permitted.then_some(compiled_frontend));
        if self.fused_context.is_none()&&permitted {
            self.fused_context=FusedLoweringContext::capture(vm,bundle.language,self.completion_permission.unwrap().1,helpers.clone(),registration).ok().map(Arc::new);
        }
        if let Some(context) = &self.fused_context {
            // Refresh only the five immutable recipe selections, not the VM's
            // namespace/module snapshot. Already running jobs retain their
            // exact selection and validate it before foreground publication.
            *context.runtime_helpers.lock().unwrap() = crate::primitives::bundle::snapshot_runtime_helpers(vm);
        }
        let _report_scope = DebugReportScope::enter(vm.debug_report());
        let _phase = DebugPhase::new(
            vm.debug_report(),
            format!("compile {}", bundle.sources.first().map_or("<unknown>", |s| s.path.to_str().unwrap_or("<path>"))),
        );
        self.refresh_worker_preparation_hints(entry_path, bundle.language, vm);
        if let Some(source)=bundle.sources.first() {
            let source_key=source.path.canonicalize().unwrap_or_else(|_|source.path.clone());
            if std::env::var_os("VYBE_PREPARATION_AUDIT").is_some() {
                eprintln!("[preparation-audit] lowering-lookup path={} permitted={} context={} ticket={} preparing={} ready={}",
                    source.path.display(), permitted, self.fused_context.is_some(),
                    self.lowering.contains_key(&source_key), self.preparing.contains_key(&source_key),
                    self.prepared_ready.contains_key(&source_key));
            }
            if let Some(ticket)=self.lowering.remove(&source_key) {
                let lowering_wait = crate::compilation_timeline::span("foreground_lowering_wait", &source_key);
                let hints=self.preparing.get(&source_key).and_then(|t|t.hint_receiver()).cloned();
                let result=if let Some(hints)=hints{loop{
                    crossbeam_channel::select!{
                        recv(ticket.result)->result=>break result,
                        recv(hints)->discovered=>match discovered{
                            Ok(hints)=>self.consume_worker_preparation_hints(std::iter::once(hints),entry_path,bundle.language,vm),
                            Err(_)=>break ticket.result.recv(),
                        },
                    }
                }}else{ticket.result.recv()};
                drop(lowering_wait);
                if let Ok(lowered)=result {
                    if lowered.valid(bundle,vm) {
                        if let Ok(compiled)=lowered.compilation {
                            if let Some(context)=&compiled.compilation.compile_dependencies {for message in context.certificate.lock().unwrap().take_diagnostics() {eprintln!("{message}");}}
                            if let Some(report)=vm.debug_report() {report.lock().unwrap().completed.push(("parallel include lowering (worker wall)".into(),lowered.elapsed));}
                            if let Some(preparing)=self.preparing.remove(&source_key){
                                if let Some(hints)=preparing.take_hints(){self.consume_worker_preparation_hints(std::iter::once(hints),entry_path,bundle.language,vm);}
                                let resolved_hints=preparing.resolved_hints();
                                // Compilation completed before its AST is sent.
                                // Keep the original ticket alive until worker
                                // handoff ends, then release its request slot.
                                let _=preparing.wait_with_early_hints(|hints|self.consume_worker_preparation_hints(std::iter::once(hints),entry_path,bundle.language,vm));
                                // Exact source/dependencies were validated above.
                                // Reuse pure child-path summaries without rescanning.
                                self.prioritize_resolved_source_hints(resolved_hints.as_ref(),&source.path,entry_path,bundle.language);
                            }
                            self.replenish_php_preparation();
                            #[cfg(any(test,v6_service_gate))] { self.completion_handoff_hits += 1; }
                            #[cfg(feature="service-verification")]if let Some(probe)=self.fused_context.as_ref().and_then(|context|context.manifest_probe.as_ref()){probe.accepted.fetch_add(1,Ordering::Relaxed);}
                            return Ok(compiled);
                        }
                    }
                    if self.fused_context.is_some()&&self.completion_permission.is_some(){self.fused_context=FusedLoweringContext::capture(vm,bundle.language,self.completion_permission.unwrap().1,helpers.clone(),registration).ok().map(Arc::new);}
                }
            }
        }
        if permitted && self.fused_context.is_some() {
            if let Some(source)=bundle.sources.first() {
                let key=source.path.canonicalize().unwrap_or_else(|_|source.path.clone());
                if let Some(lowered)=self.retired_compiled.take(&key) {
                    let valid=lowered.valid(bundle,vm);
                    if std::env::var_os("VYBE_PREPARATION_AUDIT").is_some() {
                        eprintln!("[preparation-audit] retired-compiled-lookup path={} valid={} payload_charge_bytes={} entries={}",key.display(),valid,self.retired_compiled.payload_bytes,self.retired_compiled.entries.len());
                    }
                    if valid {
                        if let Ok(compiled)=lowered.compilation {
                            if let Some(context)=&compiled.compilation.compile_dependencies {
                                for message in context.certificate.lock().unwrap().take_diagnostics() {eprintln!("{message}");}
                            }
                            self.prepare_php_source_hints(bundle,entry_path,Some(vm));
                            self.replenish_php_preparation();
                            return Ok(compiled);
                        }
                    }
                    if let Some((_,hook))=self.completion_permission {
                        self.fused_context=FusedLoweringContext::capture(vm,bundle.language,hook,helpers.clone(),registration).ok().map(Arc::new);
                    }
                }
            }
        }
        let actual_key = bundle.sources.first().map(|source|
            source.path.canonicalize().unwrap_or_else(|_| source.path.clone()));
        let ticket = actual_key.as_ref().and_then(|path| self.preparing.remove(path));
        // Discover the current parent's descendants before retiring/refilling
        // a full window. Otherwise a stale speculative descendant takes the
        // newly freed slot before the nearer source paths are prioritized.
        let current_source_hints_queued = ticket.is_none() && self.fused_context.is_some();
        if current_source_hints_queued {
            self.prepare_php_source_hints(bundle, entry_path, Some(vm));
        }
        if ticket.is_none() && std::env::var_os("VYBE_PREPARATION_AUDIT").is_some() {
            let retained=self.preparing.len();
            let ready=self.preparing.values().filter(|ticket|ticket.is_ready()).count();
            eprintln!("[preparation-audit] foreground-miss path={} retained={retained} ready={ready} pending={}",
                bundle.sources.first().map_or_else(||"<unknown>".into(),|s|s.path.display().to_string()),self.preparation_plan.len());
            if retained>0 && ready==retained {
                for path in self.preparing.keys() {eprintln!("[preparation-audit] retained-ready path={}",path.display());}
            }
        }
        if ticket.is_none() && self.fused_context.is_some() {
            let staged_only=self.lowering.keys().filter(|path|!self.preparing.contains_key(*path)).count();
            let window=4usize.saturating_sub(self.prepared_ready.len()+staged_only);
            if self.preparing.len() >= window {
                if let Some((key,retired))=self.preparation_plan.retire_completed_for_pending(
                    &mut self.preparing,actual_key.as_deref(),
                ) {
                    let hints=std::iter::from_fn(||retired.take_hints()).collect::<Vec<_>>();
                    // Neither result has been published. Cancel paired lowering
                    // ownership before admitting replacement work into this slot.
                    let lowering = self.lowering.remove(&key);
                    let completed=lowering.as_ref().map(|ticket|ticket.result.try_recv());
                    if std::env::var_os("VYBE_PREPARATION_AUDIT").is_some() {
                        // This result is already being discarded. Observe only
                        // its completed payload; never wait for speculative work
                        // or retain chunks/ASTs to collect scheduling evidence.
                        if let Some(completed) = &completed {
                            match completed {
                                Ok(lowered) => match &lowered.compilation {
                                    Ok(pending) => {
                                        let chunks = pending.code.chunks();
                                        let code_bytes: usize = chunks.iter().map(|chunk| chunk.code.len()).sum();
                                        let link_metadata_bytes = match &pending.code {
                                            LinkPayload::Raw(_) => 0,
                                            LinkPayload::Prepared(code, _) => code.metadata_bytes(),
                                        };
                                        // These are payload sizes, NOT a total
                                        // allocation bound (constants/types and
                                        // dependency certificates also own data).
                                        eprintln!("[preparation-audit] retired-lowering path={} state=compiled source_bytes={} chunks={} code_bytes={} link_metadata_bytes={}",
                                            key.display(), lowered.source.len(), chunks.len(), code_bytes, link_metadata_bytes);
                                    }
                                    Err(_) => eprintln!("[preparation-audit] retired-lowering path={} state=compile-error", key.display()),
                                },
                                Err(crossbeam_channel::TryRecvError::Empty) => eprintln!("[preparation-audit] retired-lowering path={} state=pending", key.display()),
                                Err(crossbeam_channel::TryRecvError::Disconnected) => eprintln!("[preparation-audit] retired-lowering path={} state=declined", key.display()),
                            }
                        } else {
                            eprintln!("[preparation-audit] retired-lowering path={} state=no-ticket", key.display());
                        }
                        eprintln!("[preparation-audit] retired-completed path={}",key.display());
                    }
                    if let Some(Ok(lowered))=completed {
                        let retained=self.retired_compiled.retain(key.clone(),lowered);
                        if std::env::var_os("VYBE_PREPARATION_AUDIT").is_some() {
                            eprintln!("[preparation-audit] retired-compiled-bank path={} retained={} payload_charge_bytes={} entries={}",key.display(),retained,self.retired_compiled.payload_bytes,self.retired_compiled.entries.len());
                        }
                    }
                    drop(lowering);
                    drop(retired);
                    self.consume_worker_preparation_hints(hints,entry_path,bundle.language,vm);
                    self.replenish_php_preparation();
                }
            }
        }
        // Keep the consumed ticket's window slot available for discovery
        // from this parser. Existing tickets keep other workers occupied;
        // early hints can now schedule a descendant during the parent wait.
        let mut preparation_hints_complete=false;
        let mut statement_production=None;
        let mut module = {
            let _phase = DebugPhase::current("prepare source");
            let frontend_span = crate::compilation_timeline::span("foreground_preparation_wait", bundle.sources.first().map_or(entry_path, |source| source.path.as_path()));
            let source_hints_discovered = ticket.as_ref().is_some_and(|ticket| ticket.has_source_hints());
            let resolved_hints = ticket.as_ref().and_then(|ticket| ticket.resolved_hints());
            let prepared = bundle.sources.first().and_then(|s|self.prepared_ready.remove(&s.path.canonicalize().unwrap_or_else(|_|s.path.clone()))).or_else(|| ticket.and_then(|ticket| ticket.wait_with_early_hints(|hints| {
                // Schedule descendants while this parent's parser is still
                // running, rather than blocking on the finished parent AST.
                self.consume_worker_preparation_hints(std::iter::once(hints), entry_path, bundle.language, vm);
            })));
            drop(frontend_span);
            let ready = bundle.sources.first().and_then(|source| {
                prepared
                    .filter(|prepared| {
                        // Re-read at the actual include: speculative preparation
                        // must never hide an edit or reuse an unknown dependency.
                        prepared.matches_source(&source.path, &source.code)&&prepared.matches_frontend(bundle.language)
                    })
            });
            match ready {
                Some(mut prepared) => {
                    statement_production=prepared.statement_production.take();
                    compiled_frontend=prepared.frontend_identity;
                    preparation_hints_complete=prepared.preparation_hints_complete;
                    if source_hints_discovered {
                        self.prioritize_resolved_source_hints(resolved_hints.as_ref(),&prepared.path,entry_path,bundle.language);
                        // The worker scanned these exact, validated source
                        // bytes before parsing. Retry only its unresolved
                        // operands rather than scanning the whole file again.
                        self.refresh_php_preparation_hints(&prepared.path, entry_path, bundle.language, vm);
                    } else if !current_source_hints_queued {
                        self.prepare_php_source_hints(bundle, entry_path, Some(vm));
                    }
                    if let Some(report) = vm.debug_report() {
                        report.lock().unwrap().completed.push((
                            "parallel source preparation (worker wall)".to_string(), prepared.elapsed));
                    }
                    match prepared.module {Ok(module)=>module,Err(error)=>{self.cancel_request_preparation();return Err(error)}}
                },
                None => {
                    if !current_source_hints_queued {
                        self.prepare_php_source_hints(bundle, entry_path, Some(vm));
                    }
                    let preparer=registration.preparer();
                    let result=bundle.prepared_frontend_with_sink(Some(entry_path),preparer,&mut |hints|{
                        if let Some(source)=bundle.sources.first(){self.consume_current_source_preparation_hints(source.path.clone(),hints,entry_path,bundle.language,vm);}
                    });
                    let mut prepared=match result {Ok(result)=>result,Err(error)=>{self.cancel_request_preparation();return Err(error)}};
                    preparation_hints_complete=prepared.preparation_hints_complete;
                    statement_production=prepared.statement_production.take();
                    if let (Some(hints),Some(source))=(prepared.hints.take(),bundle.sources.first()) {
                        self.consume_current_source_preparation_hints(source.path.clone(),hints,entry_path,bundle.language,vm);
                    }
                    prepared.module
                },
            }
        };
        self.replenish_php_preparation();
        if !preparation_hints_complete {self.prepare_known_php_includes(&module,bundle,entry_path);}
        self.stage_ready_lowering(vm,bundle.language,entry_path);
        fn collect_trait_names(body: &[Statement], out: &mut Vec<(String, String)>) {
            for statement in body {
                match &statement.kind {
                    StmtKind::ClassDecl { name, members, .. } => {
                        for member in members {
                            if let ClassMember::Augment(augment) = member {
                                out.push((name.clone(), augment.from.clone()));
                            }
                        }
                    }
                    StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
                        collect_trait_names(body, out)
                    }
                    _ => {}
                }
            }
        }
        let mut trait_names = Vec::new();
        collect_trait_names(&module.body, &mut trait_names);
        let has_trait_uses=!trait_names.is_empty();
        for (owner, trait_name) in trait_names {
            let key = trait_name.to_ascii_lowercase();
            if self.php_trait_declarations.keys().any(|name| name == &key || name.ends_with(&format!(".{key}"))) {
                continue;
            }
            let callbacks = vm
                .global("__php_autoload_stack")
                .and_then(|value| {
                    let Value::Object(stack) = value else {
                        return None;
                    };
                    let stack = stack.lock().unwrap();
                    let ObjectKind::Array(callbacks) = &stack.kind else {
                        return None;
                    };
                    Some(callbacks.clone())
                })
                .unwrap_or_default();
            let _phase = DebugPhase::current_lazy(|| format!("autoload trait {trait_name}"));
            let namespace = owner.rsplit_once('.').map(|(namespace, _)| namespace).unwrap_or("");
            let mut candidates = vec![trait_name.clone()];
            if !namespace.is_empty() && !trait_name.starts_with(&format!("{namespace}.")) {
                candidates.push(format!("{namespace}.{trait_name}"));
            }
            for candidate in candidates {
                let spelling = Value::String(Arc::from(candidate.replace('.', "\\")));
                for callback in &callbacks {
                    vm.invoke_callback(callback, &[spelling.clone()]);
                    if self.php_trait_declarations.keys().any(|name| name == &key || name.ends_with(&format!(".{key}"))) {
                        break;
                    }
                }
                if self.php_trait_declarations.keys().any(|name| name == &key || name.ends_with(&format!(".{key}"))) {
                    break;
                }
            }
        }
        let source_declarations = crate::preparation::source_declarations(&module);
        self.include_available_php_traits(&mut module);
        // Installed fixed providers are independently certified live VM
        // dependencies. Reusing one does not require permission to schedule
        // speculative lowering; ordinary foreground includes need the same
        // opportunity to omit duplicate bodies. Debugger, identity, ABI and
        // binding checks remain in snapshot_runtime_helpers and Certificate.
        let installed_helpers = crate::primitives::bundle::snapshot_runtime_helpers(vm);
        let track_dependencies = retain_dependencies || !installed_helpers.is_empty();
        let mut libraries=Vec::new();
        crate::preparation::bind_installed_libraries(&mut module,|name,identity| {
            let installed=library_installed(vm,name,identity);if track_dependencies {libraries.push((name.into(),identity.into(),installed));}installed
        });
        let certificate=track_dependencies.then(|| {
            let mut certificate=crate::compile_dependencies::Certificate::default();
            certificate.install_runtime_helpers(installed_helpers);
            Arc::new(std::sync::Mutex::new(certificate))
        });
        // Trait augmentation consults request-owned source declarations which
        // are not yet selectively certified. Never reuse those compiled units.
        if has_trait_uses {if let Some(certificate)=&certificate {certificate.lock().unwrap().reject();}}
        // A lowering failure can depend on missing exports/libraries. The
        // legacy error-cache interface carries no such certificate; decline
        // caching this error. Source-only parse failures remain cacheable.
        self.contextual_compile_failure=true;
        let profile_source=(bundle.language.profile_source)();
        let profile=crate::bundle::cached_language_profile(profile_source)?;
        let helpers=helpers.filter(|_|compiled_frontend==FrontendIdentity::with_registration(bundle.language,registration));
        let compiled = {
            let _phase = DebugPhase::current("lower module");
            if let Some(production)=statement_production {
                bundle.compile_prepared_statement_production(module,production,&vm.modules,false,certificate.clone(),Some(profile),helpers.as_ref(),registration.argument_selection())?
            } else if registration.capabilities.deferred_class_bodies.is_some()
                && ((crate::primitives::deferred_class_bodies::has_candidate_class_headers(&module.body)
                    && independent_lowering(&module))
                    || crate::primitives::deferred_class_bodies::has_installed_library_owned_candidate(&module.body,profile.case_sensitive)) {
                bundle.compile_prepared_certified_owned_with_emitters(module,&vm.modules,false,certificate.clone(),Some(profile),helpers.as_ref(),registration.argument_selection())?
            } else {
                bundle.compile_prepared_certified_with_emitters(&module,&vm.modules,false,certificate.clone(),Some(profile),helpers.as_ref(),registration.argument_selection())?
            }
        };
        self.contextual_compile_failure=false;
        if std::env::var_os("VYBE_TRACE_PHP_INCLUDE_HELPERS").is_some() {
            if let Some(script) = compiled.chunks.first() {
                let names: Vec<&str> = script
                    .global_inits
                    .iter()
                    .map(|init| init.name.as_str())
                    .collect();
                eprintln!(
                    "[php-include-helpers] path={} chunks={} global_inits={names:?}",
                    bundle.sources.first().map_or("<none>", |source| source.path.to_str().unwrap_or("<invalid-path>")),
                    compiled.chunks.len(),
                );
            }
        }
        Ok(PendingCompilation::raw(DynamicCompilation {
            compile_dependencies: certificate.map(|certificate| Arc::new(CompilationDependencies{certificate,libraries,profile:profile_source,profile_source:bundle.language.profile_source,language_name:bundle.language.name,frontend:compiled_frontend})),
            source_declarations: source_declarations.into(),
            required_libraries: compiled.required_libraries,
            chunks: compiled.chunks,
            host_imports: compiled.host_imports,
            component_source: compiled.component_source,
            component_binary: compiled.component_binary,
            entry_path: bundle.sources.first().map(|source| source.path.clone()),
            app_shell: compiled.app_shell,
            prepared_sources: None,
        }))
    }
}

impl JsDynamicRuntime {
    /// Run an eval'd bundle in the CALLER's VM so its definitions persist.
    ///
    /// Same assembly as [`Self::handle_dynamic_include`]: compile against the
    /// live VM's modules, append the chunks, install their globals, swap the
    /// active import tables for the child's, run, restore. That is what makes a
    /// class or function defined inside `eval` a real global rather than a
    /// value stranded in a throwaway VM's chunk table.
    fn eval_in_live_vm(
        &mut self,
        ctx: &mut HostContext,
        bundle: &Bundle,
        completion_capture: Option<&'static str>,
        preserve_existing_globals: bool,
        preserve_output_buffer_state: bool,
    ) -> Value {
        let vm = unsafe { &mut *self.vm };
        let ob_snapshot = if preserve_output_buffer_state {
            Some(vm.global(crate::primitives::io::OB_STACK).cloned())
        } else {
            None
        };
        let compiled = match bundle
            .compile_full_with_modules_and_php_entry_override_without_module_end_hooks(
                &vm.modules,
                None,
            )
        {
            Ok(compiled) => compiled,
            Err(e) => return throw_eval_error(ctx, "SyntaxError", &e),
        };

        let base_chunk_index = vm.chunks.len();
        crate::host_imports::install(vm, &compiled.host_imports);
        install_chunk_globals_with_policy(
            vm,
            &compiled.chunks,
            base_chunk_index,
            preserve_existing_globals,
        );

        let child_active_imports = compiled
            .chunks
            .first()
            .map(|chunk| chunk.imports.as_slice())
            .unwrap_or(&[]);
        let child_active_resolved_imports = match resolve_imports(vm, child_active_imports) {
            Ok(resolved) => resolved,
            Err(e) => {
                restore_eval_output_buffer_state(vm, &ob_snapshot);
                return throw_eval_error(ctx, "EvalError", &e.to_string());
            }
        };

        let result = vm.run_linked_nested(compiled.chunks, child_active_resolved_imports);


        // An uncaught throw inside eval'd code propagates to the caller —
        // checked BEFORE `result`. `raise_exception_inner` (calls.rs) defers
        // an exception whose matching handler lives outside this nested
        // execution: it stashes the REAL exception value in
        // `vm.last_exception` and returns `Err(VMError)` carrying only a
        // stringified stack trace. Matching on `result` first (as this used
        // to) took that `Err` for a genuine compile failure and wrapped the
        // stringified trace as a fake `SyntaxError`, losing the exception's
        // real type — `except ZeroDivisionError` around `eval('1/0')` could
        // never match, only a blanket `except Exception`.
        if let Some(exc) = vm.last_exception.take() {
            restore_eval_output_buffer_state(vm, &ob_snapshot);
            ctx.throw_value(exc);
            return Value::Undefined;
        }

        let run_value = match result {
            Ok(value) => value,
            Err(e) => {
                restore_eval_output_buffer_state(vm, &ob_snapshot);
                return throw_eval_error(ctx, "SyntaxError", &e.to_string());
            }
        };
        restore_eval_output_buffer_state(vm, &ob_snapshot);

        // Python's `eval` binds its expression value to a temp; read it back
        // and drop it so it does not linger as a caller global.
        match completion_capture {
            Some(name) => vm.remove_global(name).unwrap_or(Value::Undefined),
            None => run_value,
        }
    }

    fn new(caps: Capabilities) -> Self {
        Self {
            caps,
            vm: std::ptr::null_mut(),
        }
    }

    fn activate(
        &mut self,
        vm: &mut VM,
    ) -> ActiveJsRuntimeGuard {
        self.vm = vm as *mut VM;
        let previous = ACTIVE_JS_RUNTIME.with(|slot| slot.replace(Some(self as *mut _)));
        ActiveJsRuntimeGuard { previous }
    }

    fn can_dynamic_compile(&self) -> bool {
        self.caps.has(Capability::DynamicCompile)
    }

    fn handle_eval(&mut self, ctx: &mut HostContext, args: &[Value]) -> Value {
        // ⛔ INDIRECT EVAL ARRIVES WITH A RECEIVER AT ARGUMENT 0. §19.2.1
        // makes `eval` a VALUE as well as a call: `const g = eval; g("3+4")`,
        // `(0, eval)("…")`, `({ eval }).eval("…")`. Those are ordinary dynamic
        // calls, so under `ReceiverAbi::Parameter` argument 0 is the receiver
        // and the source is argument 1 — reading it at a fixed index returned
        // the receiver itself (`obj.eval("g")` printed `[object Object]`, the
        // aliased forms `undefined`).
        //
        // DIRECT `eval(...)` is emitted as a `CallHost`, which never runs
        // `call_value_inner`, so its `call_receiver_argc` stays 0 and this
        // strips nothing — the two shapes disagree about the argument list and
        // `user_args` is exactly the question "which one am I in". Verified on
        // both forms, not assumed: this is the same reasoning the numeric-stub
        // registration in `platforms/ecma/src/global.rs` already relies on.
        let args = ctx.user_args(args, 0);
        // ECMA-262 §18.2.1: if arg is not a String, return it unchanged.
        let source = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => return other.clone(),
            None => return Value::Undefined,
        };

        if !self.can_dynamic_compile() {
            return Value::Undefined;
        }

        if self.vm.is_null() {
            return Value::Undefined;
        }

        let trimmed = source.trim();
        if trimmed.is_empty() {
            return Value::Undefined;
        }

        // §19.2.1 PerformEval: parse first — invalid code throws a
        // catchable SyntaxError. The `parse_eval` hook keeps statement spans
        // in the EVAL STRING's own coordinates (the full `parse` prepends the
        // prelude, which shifted every span and silently broke the
        // completion-value split below into the no-`return` fallback). Reached
        // through the registry so this crate never names the JS language crate.
        let parse_eval = match vybe_runtime::registry::hooks("js").parse_eval {
            Some(f) => f,
            None => return throw_eval_error(ctx, "EvalError", "js eval hook not registered"),
        };
        let module = match parse_eval(trimmed) {
            Ok(module) => module,
            Err(err) => return throw_eval_error(ctx, "SyntaxError", &err),
        };

        // Directive prologue (§11.2.1): a leading "use strict" turns on the
        // early errors the (sloppy-mode) parser doesn't enforce.
        let is_strict =
            trimmed.starts_with("\"use strict\"") || trimmed.starts_with("'use strict'");
        if is_strict {
            if let Some(err) = strict_mode_early_error(trimmed, &module) {
                return throw_eval_error(ctx, "SyntaxError", &err);
            }
        }

        // §19.2.1.1: non-strict direct eval `var` declarations bind in the
        // caller's variable environment. Blank the `var` keyword of each
        // top-level declaration (offsets preserved) so it compiles as a
        // plain assignment — created as a global in the mini-VM and written
        // back below.
        let mut source_text = trimmed.to_string();
        // Names `var`-declared at eval top level — the ONLY new bindings
        // §19.2.1.1 lets a sloppy direct eval create in the caller's
        // environment (let/const/class stay inside eval's own scope).
        let mut var_names: std::collections::HashSet<String> = std::collections::HashSet::new();
        if !is_strict {
            for s in &module.body {
                // §19.2.1.3: a non-strict direct eval hoists its VAR bindings
                // into the caller's VariableEnvironment. `let`/`const` stay in
                // eval's own lexical environment.
                if let crate::ast::StmtKind::VarDecl {
                    kind: crate::ast::VarDeclKind::FunctionScoped,
                    declarations,
                } = &s.kind
                {
                    for d in declarations {
                        let mut names = std::collections::HashSet::new();
                        crate::primitives::collect_binding_pattern_names_pub(
                            &d.pattern, &mut names,
                        );
                        var_names.extend(names);
                    }
                    if let Some(off) =
                        line_col_to_offset(&source_text, s.span.start_line, s.span.start_col)
                    {
                        if source_text[off..].starts_with("var") {
                            source_text.replace_range(off..off + 3, "   ");
                        }
                    }
                }
            }
        }

        // §19.2.1.1: eval's completion value is the value of the textually
        // last statement when it is an expression statement. parse() hoists
        // function declarations to the front, so locate the last statement
        // by source span, split the SOURCE there, and wrap in a function
        // whose return yields the completion value. The mini-VM avoids the
        // outer import-table mismatch — same pattern as the Function()
        // constructor.
        let fn_name = "__vybe_eval_expr__";
        let split_at = module
            .body
            .iter()
            .max_by_key(|s| (s.span.end_line, s.span.end_col))
            .filter(|s| matches!(s.kind, crate::ast::StmtKind::Expr(_)))
            .and_then(|s| line_col_to_offset(&source_text, s.span.start_line, s.span.start_col));
        let eval_source = match split_at {
            Some(offset) => {
                let (head, tail) = source_text.split_at(offset);
                let tail = tail.trim_end().trim_end_matches(';');
                format!("function {fn_name}() {{ {head}\nreturn ({tail}); }}\n")
            }
            None => format!("function {fn_name}() {{ {source_text}\n}}\n"),
        };

        let Some(language) = languages::find_by_name("js") else {
            return Value::Undefined;
        };
        let bundle = bundle_from_source(eval_source, language, PathBuf::from("<eval>"));

        let mut eval_vm = VM::new();
        crate::primitives::platforms::register_platforms(&mut eval_vm, &self.caps);

        // Direct eval shares the caller's (global) scope: copy scalar /
        // object globals in. Function values are excluded — their
        // chunk_index refs belong to the outer VM's chunk table and would
        // be invalid if called from eval_vm.
        {
            let outer_vm = unsafe { &*self.vm };
            for (k, v) in outer_vm.globals_by_name() {
                let copy = match &v {
                    Value::Object(obj) => {
                        let o = obj.lock().unwrap();
                        !matches!(
                            o.kind,
                            ObjectKind::Function(_) | ObjectKind::HostFunction(_)
                        )
                    }
                    _ => true,
                };
                if copy {
                    eval_vm.set_global(&k, v.clone());
                }
            }
        }

        {
            let mut service =
                RuntimeCompilerService::with_capabilities(&mut eval_vm, self.caps.clone());
            if let Err(e) = service.compile_and_run_bundle(&bundle) {
                return throw_eval_error(ctx, "SyntaxError", &e);
            }
        }

        // compile_and_run_bundle ran the script chunk which stored the
        // function in eval_vm.globals via GLOBAL_SET; call it now.
        let fn_val = eval_vm
            .remove_global(fn_name)
            .or_else(|| eval_vm.remove_global(&fn_name.to_lowercase()))
            .unwrap_or(Value::Undefined);
        let result = eval_vm.invoke_callback(&fn_val, &[]);

        // §19.2.1.3: an exception thrown by the eval'd code and not caught
        // within it propagates to the *caller's* execution context. The mini
        // VM records an uncaught throw in `last_exception`; re-throw it into
        // the outer VM (via `ctx`) so the surrounding `try/catch` sees it.
        // Without this, runtime errors inside eval (const reassignment,
        // TDZ reads, etc.) would be silently swallowed.
        if let Some(exc) = eval_vm.last_exception.take() {
            ctx.throw_value(exc);
            return Value::Undefined;
        }

        // §19.2.1: assignments inside eval reach the caller's scope — write
        // non-function globals back. Shared objects were copied by Arc, so
        // in-place mutation is already visible; this covers rebinding and
        // new bindings (`eval("y = 99")`). NEW names escape ONLY when they
        // were `var`-declared (or bare-assigned) — §19.2.1.1: let/const/
        // class declared inside eval stay in eval's own environment.
        {
            let outer_vm = unsafe { &mut *self.vm };
            let let_like: std::collections::HashSet<String> = module
                .body
                .iter()
                .filter_map(|s| match &s.kind {
                    crate::ast::StmtKind::VarDecl { kind, declarations }
                        if !matches!(kind, crate::ast::VarDeclKind::FunctionScoped) =>
                    {
                        let mut names = std::collections::HashSet::new();
                        for d in declarations {
                            crate::primitives::collect_binding_pattern_names_pub(
                                &d.pattern, &mut names,
                            );
                        }
                        Some(names)
                    }
                    _ => None,
                })
                .flatten()
                .collect();
            for (k, v) in eval_vm.globals_by_name() {
                let is_function = matches!(&v, Value::Object(obj)
                    if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_)));
                if is_function {
                    continue;
                }
                let lexical = let_like.contains(&k) || let_like.contains(&k.to_lowercase());
                let pre_existing = outer_vm.has_global(&k);
                if lexical && !pre_existing {
                    continue;
                }
                outer_vm.set_global(&k, v);
            }
        }
        result
    }

    /// The universal `vybe:eval:eval` service — the compiler-as-a-service that
    /// PHP `eval`/Python `eval`/`exec` bind to (JS has its own `ecma:global:eval`
    /// via [`Self::handle_eval`]). The walker injects the source language as the
    /// 2nd argument, so this is language-generic: compile `code` as `language`
    /// in a mini-VM seeded with the caller's globals, run it, then write the
    /// globals back so definitions/assignments escape to the caller's scope.
    ///
    /// Args: `[code: String, language: String, ...extra]`. `extra` (Python's
    /// `{completion_value, namespace}` attrs) is currently advisory — the run
    /// result is returned as the completion value.
    fn handle_universal_eval(&mut self, ctx: &mut HostContext, args: &[Value]) -> Value {
        let source = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => return other.clone(),
            None => return Value::Undefined,
        };
        let language_name = match args.get(1) {
            Some(Value::String(s)) => s.to_string(),
            _ => return throw_eval_error(ctx, "EvalError", "eval: missing source language"),
        };

        if !self.can_dynamic_compile() || self.vm.is_null() {
            return Value::Undefined;
        }

        let Some(language) = languages::find_by_name(&language_name) else {
            return throw_eval_error(
                ctx,
                "EvalError",
                &format!("eval: no registered language {language_name:?}"),
            );
        };

        let attrs = args.get(2);
        let source_context = attrs.and_then(|a| object_string_prop(a, "source_context"));
        let eval_source = match source_context.as_deref() {
            Some("php_open_tag") if !source.trim_start().starts_with("<?") => {
                format!("<?php {source}")
            }
            _ => source.clone(),
        };

        // `completion_value` (attrs) — the run's value is its LAST TOP-LEVEL
        // EXPRESSION, not `null`. Language-agnostic: `EntryPoint::EvalCompletion`
        // is a common-AST rewrite (`bundle.rs`) turning the module's final
        // `StmtKind::Expr` into a `Return`, so the module's own return value —
        // `run_value` below, or `compile_and_run_bundle`'s result in the
        // namespace branch — already IS the completion value. No per-language
        // string rewrite, no separate readback.
        //
        // ⛔⛔ THIS USED TO BE A STRING REWRITE + A SECOND, ISOLATED VM.
        // `completion_value == true` rewrote the source to
        // `__vybe_eval_result__ = (source)`, ran it in a fresh `VM::new()`
        // — bypassing the live VM even when no namespace was requested — and
        // read the magic global back out of THAT VM. The isolated VM's value
        // came back as a wasm typed-null of the wrong kind (function-typed),
        // so every consumer of the result (`print`, `callable`, `==`) trapped
        // trying to `call_ref` it. PHP never hit this because PHP never sets
        // `completion_value` — its `eval()` is explicit-return
        // (`eval("1+2;")` is `NULL` in real PHP too), so the plain live-VM
        // `run_value` was already correct for PHP with zero extra machinery.
        // Python's `eval("1+2")` is implicit-last-expression (no `return`
        // possible at module scope), which is exactly the case
        // `EntryPoint::EvalCompletion` was built for and never wired to.
        let wants_value = args
            .get(2)
            .map(|a| object_bool_prop(a, "completion_value"))
            .unwrap_or(false);
        let preserve_existing_globals = attrs
            .map(|a| object_bool_prop(a, "preserve_existing_globals"))
            .unwrap_or(false);
        let preserve_output_buffer_state = attrs
            .map(|a| object_bool_prop(a, "preserve_output_buffer_state"))
            .unwrap_or(false);

        let mut bundle = bundle_from_source(eval_source, language, PathBuf::from("<eval>"));
        if wants_value {
            bundle.entry_point = EntryPoint::EvalCompletion;
        }

        // No explicit namespace dict → compile into the LIVE VM, exactly as
        // `handle_dynamic_include` does. `completion_value` alone (no
        // namespace) stays on this route too now: the AST rewrite above
        // already makes `run_value` the right answer, so there is nothing
        // left that needs the isolated VM.
        //
        // The mini-VM path below can only carry SCALARS back to the caller: it
        // deliberately skips every `ObjectKind::Function`, because a function's
        // `chunk_index` refers to the mini-VM's chunk table. A class IS a
        // constructor-function global, so `eval('class Foo {}')` could never
        // define anything — `class_exists('Foo')` stayed false, and so did
        // `function_exists` for an eval'd function. Running in the live VM
        // gives those definitions a real chunk table to belong to.
        //
        // Python's `eval(code, ns)` / `exec(code, ns)` keep the mini-VM route:
        // that form binds names FROM the dict and writes back INTO it, which is
        // isolation by definition rather than caller-global mutation — and
        // `eval(code, ns)` (unlike `exec`) may ALSO want a completion value,
        // which `bundle.entry_point` already carries into that branch too.
        let has_namespace_dict = args
            .get(2)
            .and_then(|a| object_get_prop(a, "namespace"))
            .is_some_and(|v| matches!(v, Value::Object(_)));
        if !has_namespace_dict {
            return self.eval_in_live_vm(
                ctx,
                &bundle,
                None,
                preserve_existing_globals,
                preserve_output_buffer_state,
            );
        }

        let mut eval_vm = VM::new();
        crate::primitives::platforms::register_platforms(&mut eval_vm, &self.caps);

        // Python's explicit namespace dict: `eval(code, globals[, locals])` /
        // `exec(code, ns)`. The walker forwards it as `attrs.namespace` (a
        // `dict` → `ObjectKind::Map`). When present, names bind FROM this dict
        // (not the caller's globals) and updates are written back INTO it.
        let namespace: Option<Arc<Mutex<Object>>> = args
            .get(2)
            .and_then(|a| object_get_prop(a, "namespace"))
            .and_then(|v| match v {
                Value::Object(ns) => Some(ns),
                _ => None,
            });
        // Original keys of the namespace dict — used at write-back to tell the
        // caller's bindings apart from host builtins seeded by register_all.
        let ns_keys: HashSet<String> = namespace
            .as_ref()
            .map(|ns| {
                ns_string_entries(&ns.lock().unwrap())
                    .into_iter()
                    .map(|(k, _)| k)
                    .collect()
            })
            .unwrap_or_default();

        // Seed the mini-VM's scope. With an explicit namespace, bind only its
        // entries (host builtins are already present from register_all);
        // otherwise share the caller's scalar/object globals. Function values
        // are excluded — their chunk_index refs belong to another chunk table.
        if let Some(ns) = &namespace {
            for (name, v) in ns_string_entries(&ns.lock().unwrap()) {
                eval_vm.set_global_owned(name, v);
            }
        } else {
            let outer_vm = unsafe { &*self.vm };
            for (k, v) in outer_vm.globals_by_name() {
                let copy = match &v {
                    Value::Object(obj) => {
                        let o = obj.lock().unwrap();
                        !matches!(
                            o.kind,
                            ObjectKind::Function(_) | ObjectKind::HostFunction(_)
                        )
                    }
                    _ => true,
                };
                if copy {
                    eval_vm.set_global(&k, v.clone());
                }
            }
        }

        // Pre-run keys (host builtins + seeded scope) — anything NOT here after
        // the run is a binding the eval'd code created.
        let base_keys: HashSet<String> = eval_vm.global_index.keys().cloned().collect();

        let compile_result = {
            let mut service =
                RuntimeCompilerService::with_capabilities(&mut eval_vm, self.caps.clone());
            service.compile_and_run_bundle(&bundle)
        };

        // Propagate an uncaught throw from eval'd code to the caller —
        // checked BEFORE `compile_result`, same reasoning as
        // `eval_in_live_vm` above: a deferred exception stashes its REAL
        // value in `last_exception` and returns `Err` with only a
        // stringified trace, which a naive `Err` match would misreport as a
        // `SyntaxError`.
        if let Some(exc) = eval_vm.last_exception.take() {
            ctx.throw_value(exc);
            return Value::Undefined;
        }

        let run_result = match compile_result {
            Ok(value) => value,
            Err(e) => return throw_eval_error(ctx, "SyntaxError", &e),
        };

        // Completion value: `bundle.entry_point` already turned the module's
        // last top-level expression into its `return` when `completion_value`
        // was requested (see above) — `run_result` IS the answer, no separate
        // captured-temp readback needed for any language.
        let result = run_result;

        // Definitions/assignments escape: into the explicit namespace dict if
        // one was given, otherwise the caller's globals. Skip function values.
        // For the namespace, only write updates to its own keys plus
        // brand-new bindings — never the host builtins that seeding left in
        // `base_keys`.
        match &namespace {
            Some(ns) => {
                let mut guard = ns.lock().unwrap();
                for (k, v) in eval_vm.globals_by_name() {
                    let is_function = matches!(&v, Value::Object(obj)
                        if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_)));
                    if is_function {
                        continue;
                    }
                    if ns_keys.contains(&k) || !base_keys.contains(&k) {
                        ns_string_insert(&mut guard, &k, v.clone());
                    }
                }
            }
            None => {
                let outer_vm = unsafe { &mut *self.vm };
                for (k, v) in eval_vm.globals_by_name() {
                    let is_function = matches!(&v, Value::Object(obj)
                        if matches!(obj.lock().unwrap().kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_)));
                    if !is_function {
                        outer_vm.set_global(&k, v.clone());
                    }
                }
            }
        }

        result
    }

    fn handle_function_constructor(&mut self, args: &[Value]) -> Result<Value, String> {
        if !self.can_dynamic_compile() {
            return Err("Dynamic compilation is disabled by the current capability set (missing Capability::DynamicCompile)".to_string());
        }

        let vm = unsafe { &mut *self.vm };
        let id = NEXT_JS_DYNAMIC_ID.fetch_add(1, Ordering::Relaxed);
        let symbol = format!("__vybe_dynamic_function_{id}");
        let params = if args.len() <= 1 {
            String::new()
        } else {
            args[..args.len() - 1]
                .iter()
                .map(value_to_string)
                .collect::<Vec<_>>()
                .join(",")
        };
        let body = args.last().map(value_to_string).unwrap_or_default();
        let source = format!("function {symbol}({params}) {{\n{body}\n}}\n");

        let language =
            languages::find_by_name("js").ok_or_else(|| "js language not found".to_string())?;
        let bundle = bundle_from_source(
            source,
            language,
            PathBuf::from(format!("dynamic/function_constructor_{id}.js")),
        );
        let function_global_name = symbol.to_lowercase();

        let mut function_vm = VM::new();
        crate::primitives::platforms::register_platforms(&mut function_vm, &self.caps);
        let _ = crate::adapters::register_all(&mut function_vm);
        sync_dynamic_function_globals(vm, &mut function_vm);

        {
            let mut service =
                RuntimeCompilerService::with_capabilities(&mut function_vm, self.caps.clone());
            service.compile_and_run_bundle(&bundle)?;
        }

        let function = function_vm
            .remove_global(&function_global_name)
            .ok_or_else(|| format!("dynamic Function did not publish {function_global_name}"))?;
        let length = dynamic_function_length(&function);
        let state = Box::new(ConstructedJsFunction {
            caps: self.caps.clone(),
            vm: function_vm,
            function,
        });
        let state_ptr = Box::into_raw(state);
        CONSTRUCTED_JS_FUNCTIONS.with(|slot| {
            slot.borrow_mut().insert(id, state_ptr);
        });

        let host_module = "vybe:js.dynamic";
        let host_name = symbol.clone();
        let dynamic_id = id;
        // ⛔ A `Function(...)`-CONSTRUCTED FUNCTION HAS NO RECEIVER
        // PARAMETER. Its declared parameters are the ones the caller named
        // (`Function("x","return x*2")`), so a receiver in front of them binds
        // `x` and the body answers NaN. Declaring the type receiverless drops
        // the slot at the call instead of making the body compensate.
        vm.register_free_fn(
            host_module,
            &host_name,
            Box::new(move |ctx, args| {
                CONSTRUCTED_JS_FUNCTIONS.with(|slot| {
                    let Some(&state_ptr) = slot.borrow().get(&dynamic_id) else {
                        return throw_dynamic_compile_error(
                            ctx,
                            format!("dynamic Function backing state missing for id {dynamic_id}"),
                        );
                    };
                    let state = unsafe { &mut *state_ptr };

                    ACTIVE_JS_RUNTIME.with(|runtime_slot| {
                        if let Some(runtime_ptr) = *runtime_slot.borrow() {
                            let runtime = unsafe { &*runtime_ptr };
                            if !runtime.vm.is_null() {
                                let source_vm = unsafe { &*runtime.vm };
                                sync_dynamic_function_globals(source_vm, &mut state.vm);
                            }
                        }
                    });

                    // §10.2.1: the body's `this` is argument 0 of the invoke
                    // below, and the nested VM's prepend reads it from the host
                    // receiver channel — `new Function("v","this.v = v")` used
                    // as a constructor has no other way to reach its instance.
                    let previous_receiver = state.vm.set_host_receiver(ctx.current_js_this());
                    ensure_js_runtime_registered(&mut state.vm);
                    let mut nested_runtime = JsDynamicRuntime::new(state.caps.clone());
                    // None from an unopted caller keeps its legacy compile route.
                    let _guard = nested_runtime.activate(&mut state.vm);
                    // ⛔ THE RECEIVER IS THIS WRAPPER'S ARGUMENT 0, NOT THE DYNAMIC
                    // FUNCTION'S. Under `ReceiverAbi::Parameter` the call site
                    // hands a host callee its receiver at argument 0, so the
                    // compiled body must not see it: without this,
                    // `Function("x","return x*2")(9)` is `NaN` because `x`
                    // binds the receiver.
                    //
                    // ⛔ THIS WAS WRONG ONCE, FOR A REASON WORTH KEEPING. The
                    // strip is only sound if EVERY call site actually pushes a
                    // receiver, and `receiver_argc()` cannot tell you that — it
                    // is computed from the module ABI and from who dispatched,
                    // never from the argument list. The two construction forms
                    // used to disagree: a function from `Function(...)` was
                    // called through the ordinary receiver-carrying path, while
                    // one from `new Function(...)` got a `Function` type hint
                    // and was lowered as a MULTICAST DELEGATE, whose ladder
                    // passes handler arguments only. Stripping on a number that
                    // read 1 for both fixed 3 tests and broke 6 (corpus-
                    // measured, net -4), so it was withdrawn.
                    //
                    // It is correct now because the disagreement is gone at the
                    // source: the delegate lowering is no longer selected under
                    // `Parameter` (see `calls.rs`), so both forms pass a
                    // receiver and `receiver_argc()` is finally describing the
                    // list it claims to describe. Fix the emitter, THEN trust
                    // the number — never the other way round.
                    let inner_args = ctx.user_args(args, 0);
                    let result = match state.vm.invoke(&state.function, inner_args) {
                        Ok(value) => value,
                        Err(err) => throw_dynamic_compile_error(ctx, err.to_string()),
                    };

                    state.vm.set_host_receiver(previous_receiver);

                    result
                })
            }),
        );

        let host_idx = *vm
            .host_registry
            .get(&(host_module.to_string(), host_name.clone()))
            .ok_or_else(|| format!("dynamic Function host wrapper missing for {host_name}"))?;
        Ok(dynamic_host_function_value(
            vm,
            host_module,
            &host_name,
            host_idx,
            length,
        ))
    }
}

impl Drop for ActivePhpRuntimeGuard {
    fn drop(&mut self) {
        ACTIVE_PHP_RUNTIME.with(|slot| {
            slot.replace(self.previous.take());
        });
    }
}

impl Drop for ActiveJsRuntimeGuard {
    fn drop(&mut self) {
        ACTIVE_JS_RUNTIME.with(|slot| {
            slot.replace(self.previous.take());
        });
    }
}

fn ensure_php_runtime_registered(vm: &mut VM) {
    let compact_key = ("vybe:php".to_string(), "compact".to_string());
    if !vm.host_registry.contains_key(&compact_key) {
        vm.register_host_fn(
            "vybe:php",
            "compact",
            Box::new(|_, args| {
                fn names(value: &Value, output: &mut Vec<String>) {
                    match value {
                        Value::String(name) => output.push(name.to_string()),
                        Value::Object(object) => {
                            let values = match &object.lock().unwrap().kind {
                                ObjectKind::Array(values) => values.clone(),
                                ObjectKind::Map(values) => values.values().cloned().collect(),
                                _ => Vec::new(),
                            };
                            for value in values {
                                names(&value, output);
                            }
                        }
                        _ => {}
                    }
                }
                ACTIVE_PHP_RUNTIME.with(|active| {
                    let Some(pointer) = *active.borrow() else {
                        return Value::Null;
                    };
                    let runtime = unsafe { &mut *pointer };
                    if runtime.vm.is_null() {
                        return Value::Null;
                    }
                    let vm = unsafe { &mut *runtime.vm };
                    let mut variables = HashMap::new();
                    if let Some((chunk, locals)) = vm.current_source_scope() {
                        if chunk != "<script>" {
                            for (_, name, value) in locals {
                                variables.insert(name.trim_start_matches('$').to_string(), value);
                            }
                        } else {
                            let scope = runtime.include_scope.clone().unwrap_or_else(|| {
                                let Some(Value::Object(global)) = vm.global("globalThis") else {
                                    return Value::Null;
                                };
                                global
                                    .lock()
                                    .unwrap()
                                    .properties
                                    .get("__vybe_php_globals")
                                    .cloned()
                                    .unwrap_or(Value::Null)
                            });
                            if let Value::Object(scope) = scope {
                                variables.extend(
                                    scope
                                        .lock()
                                        .unwrap()
                                        .properties
                                        .iter()
                                        .map(|(key, value)| (key.clone(), value.clone())),
                                );
                            }
                        }
                    }
                    let mut requested = Vec::new();
                    for arg in args {
                        names(arg, &mut requested);
                    }
                    let mut result = Object::new();
                    result.kind = ObjectKind::Map(Default::default());
                    if let ObjectKind::Map(entries) = &mut result.kind {
                        for name in requested {
                            if let Some(value) = variables.get(&name) {
                                let value =
                                    if let Some(deref) = vm.global("__vybe_autoderef").cloned() {
                                        vm.invoke_callback(&deref, &[value.clone()])
                                    } else {
                                        value.clone()
                                    };
                                entries.insert(Value::String(Arc::from(name)), value);
                            }
                        }
                    }
                    Value::Object(vybe_runtime::heap::alloc(result))
                })
            }),
        );
    }
    let scope_key = ("vybe:php".to_string(), "include_scope".to_string());
    if !vm.host_registry.contains_key(&scope_key) {
        vm.register_host_fn("vybe:php", "include_scope", Box::new(|_, _| php_include_scope()),);
    }
    let include_key = ("vybe:php".to_string(), "dynamic_include".to_string());
    if !vm.host_registry.contains_key(&include_key) {
        vm.register_host_fn(
            "vybe:php",
            "dynamic_include",
            Box::new(|ctx, args| {
                ACTIVE_PHP_RUNTIME.with(|slot| {
                    let Some(runtime_ptr) = *slot.borrow() else {
                        return throw_dynamic_compile_error(
                            ctx,
                            "include/require ran with no active PHP runtime".to_string(),
                        );
                    };
                    let runtime = unsafe { &mut *runtime_ptr };
                    match runtime.handle_dynamic_include(args) {
                        Ok(value) => value,
                        Err(message) => throw_php_include_error(ctx, message),
                    }
                })
            }),
        );
    }

    let function_key = ("vybe:php".to_string(), "global_function".to_string());
    if !vm.host_registry.contains_key(&function_key) {
        vm.register_host_fn(
            "vybe:php",
            "global_function",
            Box::new(|ctx, args| {
                let raw_name = args.first().map(value_to_string).unwrap_or_default();
                ACTIVE_PHP_RUNTIME.with(|slot| {
                    let Some(runtime_ptr) = *slot.borrow() else {
                        return throw_dynamic_compile_error(
                            ctx,
                            "PHP global function lookup ran with no active runtime".to_string(),
                        );
                    };
                    let runtime = unsafe { &mut *runtime_ptr };
                    if runtime.vm.is_null() {
                        return Value::Null;
                    }
                    let vm = unsafe { &mut *runtime.vm };
                    let trimmed = raw_name.trim_start_matches('\\');
                    let dotted = trimmed.replace('\\', ".");
                    let candidates = [
                        raw_name.clone(),
                        trimmed.to_string(),
                        dotted.clone(),
                        raw_name.to_lowercase(),
                        trimmed.to_lowercase(),
                        dotted.to_lowercase(),
                        trimmed.strip_prefix("__php_fn_").unwrap_or(trimmed).to_lowercase(),
                    ];
                    for name in candidates {
                        if let Some(value) = vm.global(&name) {
                            if !matches!(value, Value::Null | Value::Undefined) {
                                return value.clone();
                            }
                        }
                    }
                    if let Some(argc) = args.get(1) {
                        let name = trimmed.strip_prefix("__php_fn_").unwrap_or(trimmed);
                        return match runtime.builtin_callable(name, argc.as_f64() as usize) {
                            Ok(value) => value,
                            Err(message) => throw_dynamic_compile_error(ctx, message),
                        };
                    }
                    Value::Null
                })
            }),
        );
    }
}

fn throw_php_include_error(ctx: &mut HostContext, message: String) -> Value {
    // A failed PHP require is catchable as Error/Throwable, not as a string.
    // Allocate through the installed PHP class factory so its declared type
    // identity and ordinary exception methods remain intact.
    let constructor = ctx.get_global("Error");
    if !matches!(constructor, Value::Null | Value::Undefined) {
        match ctx.try_invoke_with_receiver(
            &constructor,
            Value::Undefined,
            &[Value::String(message.clone().into())],
        ) {
            Ok(error @ Value::Object(_)) => {
                ctx.throw_value(error);
                return Value::Null;
            }
            Err(error) => {
                ctx.throw_value(error);
                return Value::Null;
            }
            Ok(_) => {}
        }
    }
    throw_dynamic_compile_error(ctx, message)
}

fn throw_dynamic_compile_error(ctx: &mut HostContext, message: String) -> Value {
    ctx.throw_value(Value::String(message.into()));
    Value::Null
}

/// Read a boolean property off a `Value::Object` (used for the Python eval
/// attrs object `{completion_value, namespace}` the walker injects). Handles
/// both object shapes: `Ordinary` (properties) and `Map` (dict).
fn object_bool_prop(v: &Value, key: &str) -> bool {
    if let Value::Object(obj) = v {
        let o = obj.lock().unwrap();
        let found = match &o.kind {
            ObjectKind::Map(m) => m.get(&Value::String(key.into())).cloned(),
            _ => o.properties.get(key).cloned(),
        };
        return matches!(found, Some(Value::Bool(true)));
    }
    false
}

fn object_string_prop(v: &Value, key: &str) -> Option<String> {
    object_get_prop(v, key).and_then(|value| match value {
        Value::String(s) => Some(s.to_string()),
        _ => None,
    })
}

fn restore_eval_output_buffer_state(vm: &mut VM, snapshot: &Option<Option<Value>>) {
    let Some(snapshot) = snapshot else {
        return;
    };
    match snapshot {
        Some(value) => vm.set_global(crate::primitives::io::OB_STACK, value.clone()),
        None => {
            vm.remove_global(crate::primitives::io::OB_STACK);
        }
    }
}

/// Read a property off a `Value::Object` regardless of `Map`/`Ordinary` shape.
fn object_get_prop(v: &Value, key: &str) -> Option<Value> {
    if let Value::Object(obj) = v {
        let o = obj.lock().unwrap();
        return match &o.kind {
            ObjectKind::Map(m) => m.get(&Value::String(key.into())).cloned(),
            _ => o.properties.get(key).cloned(),
        };
    }
    None
}

/// String-keyed entries of a namespace object (Python `dict`), reading both the
/// `Map` and `Ordinary` shapes.
fn ns_string_entries(obj: &Object) -> Vec<(String, Value)> {
    match &obj.kind {
        ObjectKind::Map(m) => m
            .iter()
            .filter_map(|(k, v)| match k {
                Value::String(s) => Some((s.to_string(), v.clone())),
                _ => None,
            })
            .collect(),
        _ => obj
            .properties
            .iter()
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect(),
    }
}

/// Insert a string-keyed binding into a namespace object, matching its shape.
fn ns_string_insert(obj: &mut Object, key: &str, val: Value) {
    match &mut obj.kind {
        ObjectKind::Map(m) => {
            m.insert(Value::String(key.into()), val);
        }
        _ => {
            obj.properties.insert(key.to_string(), val);
        }
    }
}

/// Throw a JS-shaped error object (same stamp `vybe_host`'s error machinery
/// uses) so `catch (e) { e instanceof SyntaxError }` works on eval failures.
fn throw_eval_error(ctx: &mut HostContext, kind: &str, message: &str) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from(kind)));
    obj.properties
        .insert("__exception_type".into(), Value::String(Arc::from(kind)));
    obj.properties
        .insert("name".into(), Value::String(Arc::from(kind)));
    obj.properties
        .insert("message".into(), Value::String(Arc::from(message)));
    obj.properties.insert(
        "stack".into(),
        Value::String(Arc::from(format!("{kind}: {message}").as_str())),
    );
    let chain = Object::new_array(vec![
        Value::String(Arc::from(kind)),
        Value::String(Arc::from("Error")),
    ]);
    obj.properties.insert(
        "__types".into(),
        Value::Object(vybe_runtime::heap::alloc(chain)),
    );
    // ⛔ LINK THE PROTOTYPE. Without this the error carries the right NAME and
    // the right `__types` chain and is still not an instance of anything:
    // `eval("if (") ` threw an object for which `e instanceof SyntaxError` and
    // even `e instanceof Error` were both `false`.
    //
    // `platforms/ecma`'s `error.rs` states the rule on `new_error_flat` —
    // *"every throw-site constructor must use `new_error`; an unlinked error is
    // the two-populations bug"* — and this is a throw site that had grown its
    // own error object instead. It only looked correct because `js_instanceof`
    // used to answer from the `__type` stamp above.
    //
    // Resolved the same way `link_error_prototype` does, through the per-VM
    // `__ctor_<Kind>` anchor rather than a shared static: the error prototypes
    // belong to the running VM, and reaching into `platforms/ecma` from the
    // compiler would invert the layering. A language whose profile wires no
    // error constructors simply finds nothing and keeps the stamps.
    if let Value::Object(ctor) = ctx.get_global(&format!("__ctor_{kind}")) {
        let proto = ctor.lock().unwrap().properties.get("prototype").cloned();
        if let Some(proto @ Value::Object(_)) = proto {
            obj.properties.insert("__proto__".into(), proto);
            // §20.5.3.2: with a chain in place `name` resolves THROUGH the
            // prototype — `new SyntaxError("x").hasOwnProperty("name")` is
            // false in node. Leaving the own stamp made this the one throw
            // site out of nineteen that reported `ownName=true`.
            // Only when the link actually happened: a language that wires no
            // error constructors keeps the stamp, or it loses its name.
            obj.properties.shift_remove("name");
        }
    }
    ctx.throw_value(Value::Object(vybe_runtime::heap::alloc(obj)));
    Value::Undefined
}

/// Strict-mode early errors (§12.9.4.1 legacy octals, §13.1.1 reserved
/// words as bindings) that the sloppy-mode parser accepts.
fn strict_mode_early_error(src: &str, module: &crate::ast::Module) -> Option<String> {
    // Legacy octal literal: `0` followed by octal digits, outside strings,
    // not a 0x/0o/0b prefix and not part of a longer number/identifier.
    let bytes = src.as_bytes();
    let mut in_string: Option<u8> = None;
    let mut prev: u8 = b' ';
    let mut i = 0;
    while i < bytes.len() {
        let c = bytes[i];
        if let Some(q) = in_string {
            if c == q && prev != b'\\' {
                in_string = None;
            }
        } else if c == b'"' || c == b'\'' || c == b'`' {
            in_string = Some(c);
        } else if c == b'0'
            && !prev.is_ascii_alphanumeric()
            && prev != b'_'
            && prev != b'$'
            && prev != b'.'
        {
            let next = bytes.get(i + 1).copied().unwrap_or(b' ');
            if (b'0'..=b'7').contains(&next) {
                return Some("Octal literals are not allowed in strict mode".to_string());
            }
        }
        prev = c;
        i += 1;
    }

    const RESERVED: [&str; 9] = [
        "implements",
        "interface",
        "package",
        "private",
        "protected",
        "public",
        "static",
        "let",
        "yield",
    ];
    for s in &module.body {
        if let crate::ast::StmtKind::VarDecl { declarations, .. } = &s.kind {
            for d in declarations {
                if let crate::ast::BindingPattern::Ident(name) = &d.pattern {
                    if RESERVED.contains(&name.as_str()) {
                        return Some(format!("Unexpected strict mode reserved word: {name}"));
                    }
                }
            }
        }
    }
    None
}

/// Byte offset of a 0-based (line, col) position in `src`.
fn line_col_to_offset(src: &str, line: u32, col: u32) -> Option<usize> {
    let mut current = 0usize;
    for (i, l) in src.split('\n').enumerate() {
        if i as u32 == line {
            let col = col as usize;
            return (col <= l.len()).then_some(current + col);
        }
        current += l.len() + 1;
    }
    None
}

fn ensure_js_runtime_registered(vm: &mut VM) {
    let key = ("vybe:js".to_string(), "function_constructor".to_string());
    if vm.host_registry.contains_key(&key) {
        return;
    }

    vm.register_free_fn(
        "vybe:js",
        "function_constructor",
        Box::new(|ctx, args| {
            ACTIVE_JS_RUNTIME.with(|slot| {
                let Some(runtime_ptr) = *slot.borrow() else {
                    return throw_dynamic_compile_error(
                        ctx,
                        "JS dynamic runtime is not active for Function construction".to_string(),
                    );
                };
                let runtime = unsafe { &mut *runtime_ptr };
                match runtime.handle_function_constructor(args) {
                    Ok(value) => value,
                    Err(err) => throw_dynamic_compile_error(ctx, err),
                }
            })
        }),
    );

    vm.register_free_fn(
        "ecma:global",
        "eval",
        Box::new(|ctx, args| {
            ACTIVE_JS_RUNTIME.with(|slot| {
                let Some(runtime_ptr) = *slot.borrow() else {
                    return throw_dynamic_compile_error(
                        ctx,
                        "JS dynamic runtime is not active for eval".to_string(),
                    );
                };
                let runtime = unsafe { &mut *runtime_ptr };
                runtime.handle_eval(ctx, args)
            })
        }),
    );

    // The universal `vybe:eval:eval` compiler-as-a-service — what PHP `eval`
    // and Python `eval`/`exec` bind to (language injected as arg 1). Routed
    // through the always-active JS runtime, which carries the outer VM + caps.
    vm.register_free_fn(
        "vybe:eval",
        "eval",
        Box::new(|ctx, args| {
            ACTIVE_JS_RUNTIME.with(|slot| {
                let Some(runtime_ptr) = *slot.borrow() else {
                    return throw_dynamic_compile_error(
                        ctx,
                        "eval service is not active (no dynamic runtime)".to_string(),
                    );
                };
                let runtime = unsafe { &mut *runtime_ptr };
                runtime.handle_universal_eval(ctx, args)
            })
        }),
    );

    // §19.2.1: eval is also a *value* (`const e = eval; e("…")` — indirect
    // eval). Expose the registered host fn as a callable global, same shape
    // vybe_host's namespace machinery uses.
    if let Some(&idx) = vm
        .host_registry
        .get(&("ecma:global".to_string(), "eval".to_string()))
    {
        let mut obj = Object::new();
        obj.properties
            .insert("name".into(), Value::String(Arc::from("eval")));
        obj.kind = ObjectKind::HostFunction(idx);
        vm.set_global("eval", Value::Object(vybe_runtime::heap::alloc(obj)));
    }
}

// Runtime include bases are canonical active-file identities; an absolute
// target retains its spelling (including parser-produced __DIR__ literals).
pub(crate) fn resolve_prepared_php_include_path(entry: &Path, caller: &Path, target: &str, fused: bool) -> PathBuf {
    if !fused { return resolve_php_include_path(entry, caller, target); }
    let entry = entry.canonicalize().unwrap_or_else(|_|entry.to_path_buf());
    let caller = caller.canonicalize().unwrap_or_else(|_|caller.to_path_buf());
    resolve_php_include_path(&entry, &caller, target)
}

fn resolve_php_include_path(entry_path: &Path, caller_path: &Path, target: &str) -> PathBuf {
    let target_path = PathBuf::from(target);
    if target_path.is_absolute() {
        return target_path;
    }
    let base_path = if is_explicit_relative_php_include(&target_path) {
        caller_path
    } else {
        entry_path
    };
    base_path
        .parent()
        .map(|parent| parent.join(target_path.clone()))
        .unwrap_or(target_path)
}

fn is_explicit_relative_php_include(path: &Path) -> bool {
    let raw = path.to_string_lossy();
    raw == "."
        || raw == ".."
        || raw.starts_with("./")
        || raw.starts_with("../")
        || raw.starts_with(".\\")
        || raw.starts_with("..\\")
}

fn value_to_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        _ => format!("{value}"),
    }
}

fn global_constructor_prototype(vm: &VM, name: &str) -> Option<Value> {
    let Value::Object(ctor) = vm.global(name)?.clone() else {
        return None;
    };
    let ctor = ctor.lock().unwrap();
    ctor.properties.get("prototype").cloned()
}

fn dynamic_function_length(value: &Value) -> f64 {
    let Value::Object(function_obj) = value else {
        return 0.0;
    };
    let function_obj = function_obj.lock().unwrap();
    if let Some(Value::F64(length)) = function_obj.properties.get("length") {
        return *length;
    }
    match &function_obj.kind {
        ObjectKind::Function(function) => function.arity as f64,
        _ => 0.0,
    }
}

fn dynamic_host_function_value(
    vm: &VM,
    module: &str,
    name: &str,
    host_idx: usize,
    length: f64,
) -> Value {
    let Some(function_proto) = global_constructor_prototype(vm, "Function") else {
        return Value::Null;
    };
    let object_proto = global_constructor_prototype(vm, "Object");

    let mut function_obj = Object::new();
    function_obj.kind = ObjectKind::HostFunction(host_idx);
    function_obj
        .properties
        .insert("__host_module".into(), Value::String(Arc::from(module)));
    function_obj
        .properties
        .insert("__host_name".into(), Value::String(Arc::from(name)));
    function_obj
        .properties
        .insert("__host_idx".into(), Value::F64(host_idx as f64));
    function_obj
        .properties
        .insert("name".into(), Value::String(Arc::from("anonymous")));
    function_obj
        .properties
        .insert("length".into(), Value::F64(length));
    function_obj
        .properties
        .insert("__proto__".into(), function_proto.clone());

    let function_value = Value::Object(vybe_runtime::heap::alloc(function_obj));
    let mut prototype = Object::new();
    prototype
        .properties
        .insert("constructor".into(), function_value.clone());
    if let Some(object_proto) = object_proto {
        prototype
            .properties
            .insert("__proto__".into(), object_proto);
    }
    if let Value::Object(function_obj) = &function_value {
        function_obj.lock().unwrap().properties.insert(
            "prototype".into(),
            Value::Object(vybe_runtime::heap::alloc(prototype)),
        );
    }

    function_value
}

fn sync_dynamic_function_globals(source: &VM, target: &mut VM) {
    for (name, value) in source.globals_by_name() {
        if is_shared_dynamic_global(&name, &value) {
            target.set_global(&name, value);
        }
    }
}

fn is_shared_dynamic_global(name: &str, value: &Value) -> bool {
    if !name.eq_ignore_ascii_case("globalThis") {
        return false;
    }

    match value {
        Value::Object(object) => {
            let object = object.lock().unwrap();
            !matches!(
                object.kind,
                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
            )
        }
        _ => false,
    }
}

/// Thin wrapper over the VM's single import-resolution policy
/// (`VM::resolve_import_target`) — this used to be a third hand-rolled copy
/// of the loop, drifted to a raw `host_registry` lookup that bypassed
/// Module Record exports.
/// Resolve a child module's imports against the live VM.
///
/// A free global (`env` / `*`) that is not defined YET is not an error: a
/// runtime `include`/`require` — or an `eval` — may define it a moment from
/// now, and resolving eagerly to a link failure is what made a function
/// defined by an included file unreachable from its caller. Those names fall
/// back to [`ImportTarget::StdlibRedirect`], which holds the NAME and looks it
/// up in `vm.globals` at CALL time; `install_chunk_globals` publishes exactly
/// that key when the include runs. Still missing when called, the VM reports
/// it then — which is where PHP reports an undefined function too.
///
/// The fallback is deliberately limited to the embedder-provided namespaces.
/// A `wasi:*` or `vybe:*` import that does not resolve is a genuinely missing
/// host function and must still fail loudly at link time.
fn resolve_imports(vm: &VM, imports: &[Import]) -> Result<Vec<ImportTarget>, String> {
    imports
        .iter()
        .map(
            |import| match vm.resolve_import_target(&import.module, &import.name) {
                Ok(target) => Ok(target),
                Err(_) if is_free_global_module(&import.module) => {
                    Ok(ImportTarget::StdlibRedirect(import.name.clone()))
                }
                Err(e) => Err(e.to_string()),
            },
        )
        .collect()
}

/// The namespaces `globals::declare_free_globals` and the bundler treat as
/// embedder-provided. Kept in sync with `bundle.rs`, which groups `"*"`,
/// `"env"` and `"wasm:string-constants"` the same way.
fn is_free_global_module(module: &str) -> bool {
    matches!(module, "env" | "*")
}

#[cfg(test)]
mod tests {
    #[test]
    fn frontend_identity_separates_each_compiler_recipe_option() {
        use vybe_runtime::registry::CompilerOptimizations;
        vybe_language_php::register();
        let language = crate::languages::find_by_name("php").unwrap();
        let baseline = CompilerOptimizations::default();
        let identity = |options| super::FrontendIdentity::with_options_identity(language, [0;10], options);
        let before = identity(baseline);
        let variants = [
            CompilerOptimizations {symbol_membership_index:true, ..baseline},
            CompilerOptimizations {user_namespace_cache:true, ..baseline},
            CompilerOptimizations {completion_worker_lowering:true, ..baseline},
            CompilerOptimizations {compact_source_lines:true, ..baseline},
            CompilerOptimizations {array_call_tail_identity:true, ..baseline},
            CompilerOptimizations {interpolation_call_tail_identity:true, ..baseline},
            CompilerOptimizations {fused_closure_analysis:true, ..baseline},
        ];
        for (index, options) in variants.iter().enumerate() {
            assert_ne!(before, identity(*options), "recipe option {index} must invalidate cached compilation");
            for other in &variants[..index] {
                assert_ne!(identity(*other), identity(*options));
            }
        }
        assert_eq!(before, identity(baseline));
    }

    #[test]
    fn worker_library_snapshot_preserves_exact_and_lowercase_shadowing() {
        use std::collections::HashMap;
        let names = ["foo", "Foo", "FOO", "bar", "Bar", "BAR", "İ", "i\u{307}", "Σ", "σ", "ς", "K", "k", ""];
        let lookup = |bindings: &HashMap<String, Option<String>>, name: &str, identity: &str| {
            bindings.get(name).or_else(|| bindings.get(&name.to_lowercase()))
                .is_some_and(|value| value.as_deref() == Some(identity))
        };
        for seed in 0usize..1024 {
            let bindings: Vec<_> = names.iter().enumerate().map(|(i, name)| {
                ((*name).to_owned(), if seed.rotate_left(i as u32) % 7 < 2 {
                    Some(format!("identity{}", (seed + i) % 3))
                } else { None })
            }).collect();
            let original: HashMap<_, _> = bindings.iter().cloned().collect();
            let compact = super::snapshot_library_bindings(bindings.iter().map(|(name, identity)| (name.as_str(), identity.clone())));
            for name in names.into_iter().chain(["fOo", "bAr", "unknown", "UNKNOWN"]) {
                for identity in ["identity0", "identity1", "identity2", ""] {
                    assert_eq!(lookup(&original, name, identity), lookup(&compact, name, identity), "seed={seed}, name={name:?}");
                }
            }
        }
    }

    #[test]
    fn worker_library_snapshot_does_not_retain_unrelated_ascii_globals() {
        let bindings = vec![
            ("library", Some("owner".to_owned())),
            ("Library", None),
            ("ordinary_global", None),
            ("A_Long_Unrelated_Constant_Binding", None),
            ("Σordinary_global", None),
        ];
        let compact = super::snapshot_library_bindings(bindings.into_iter());
        assert_eq!(compact.len(), 3);
        assert_eq!(compact.get("library"), Some(&Some("owner".to_owned())));
        assert_eq!(compact.get("Library"), Some(&None));
        assert_eq!(compact.get("Σordinary_global"), Some(&None));
        assert!(!compact.contains_key("ordinary_global"));
        assert!(!compact.contains_key("A_Long_Unrelated_Constant_Binding"));
    }

    #[test]
    fn speculative_link_admission_is_bounded_and_released_on_drop() {
        use super::{LinkPermit, LINK_OWNER_BUDGET, LINK_REQUEST_BUDGET};
        use std::sync::atomic::{AtomicUsize, Ordering};
        let budget = Arc::new(AtomicUsize::new(0));
        let first = LinkPermit::reserve(&budget).unwrap();
        let mut second = LinkPermit::reserve(&budget).unwrap();
        assert_eq!(budget.load(Ordering::Relaxed), LINK_REQUEST_BUDGET);
        assert!(LinkPermit::reserve(&budget).is_none());
        second.retain(123);
        assert_eq!(budget.load(Ordering::Relaxed), LINK_OWNER_BUDGET + 123);
        assert!(LinkPermit::reserve(&budget).is_none());
        drop(first);
        let third = LinkPermit::reserve(&budget).unwrap();
        drop(second);
        drop(third);
        assert_eq!(budget.load(Ordering::Relaxed), 0);
        // Requests share no admission state, even when an old job is alive.
        let other = Arc::new(AtomicUsize::new(0));
        let permit = LinkPermit::reserve(&budget).unwrap();
        assert_eq!(other.load(Ordering::Relaxed), 0);
        drop(permit);
    }

    #[test]
    fn speculative_link_owner_keeps_public_mutation_raw() {
        use super::{DynamicCompilation, PendingCompilation, LinkPayload, LinkPermit};
        use std::sync::atomic::{AtomicUsize, Ordering};
        let mut chunk = vybe_runtime::Chunk::new("handoff");
        chunk.emit_f64_const(41.0, 1);
        chunk.emit_op(vybe_runtime::Op::RETURN, 1);
        let expected = chunk.code.clone();
        let make = || DynamicCompilation {
            compile_dependencies: None, source_declarations: Arc::from([]),
            required_libraries: vec![], chunks: vec![chunk.clone()],
            host_imports: Default::default(), component_source: false,
            component_binary: None, entry_path: None, prepared_sources: None, app_shell: None,
        };
        let budget = Arc::new(AtomicUsize::new(0));
        let pending = PendingCompilation::prepare(make(), &budget);
        assert!(pending.compilation.chunks.is_empty());
        assert!(matches!(&pending.code, LinkPayload::Prepared(..)));
        assert_eq!(pending.code.chunks()[0].code, expected);
        assert!(budget.load(Ordering::Relaxed) > 0);
        let mut raw = pending.into_raw();
        assert_eq!(budget.load(Ordering::Relaxed), 0);
        assert_eq!(raw.chunks[0].code, expected);
        raw.chunks[0].code.clear(); // public mutation can never retain trusted metadata
        let first = LinkPermit::reserve(&budget).unwrap();
        let second = LinkPermit::reserve(&budget).unwrap();
        let fallback = PendingCompilation::prepare(make(), &budget);
        assert!(matches!(&fallback.code, LinkPayload::Raw(..)));
        assert_eq!(fallback.into_raw().chunks[0].code, expected);
        drop(first); drop(second);
        assert_eq!(budget.load(Ordering::Relaxed), 0);
        drop(PendingCompilation::prepare(make(), &budget));
        assert_eq!(budget.load(Ordering::Relaxed), 0);
    }


    #[test]
fn immediate_ast_lookahead_skips_deferred_bodies_and_default_is_exact() {
    use vybe_ast::*;
    fn expression(name:&str)->Statement {Statement::new(StmtKind::Expr(Expression::ident(name)))}
    let function=Statement::new(StmtKind::FunctionDecl {
        name:"later".into(),params:vec![],return_type:None,body:vec![expression("cold")],
        modifiers:Default::default(),handles:vec![],is_async:false,is_generator:false,is_sub:false,
    });
    let class=Statement::new(StmtKind::ClassDecl {name:"Deferred".into(),parents:vec![],interfaces:vec![],
        members:vec![ClassMember::Method(Box::new(function.clone()))],modifiers:Default::default(),decorators:vec![]});
    let root=Statement::new(StmtKind::NamespaceDecl {name:"N".into(),body:vec![
        Statement::new(StmtKind::Block(vec![function,class,expression("hot")])),
        Statement::new(StmtKind::If{cond:Expression::ident("condition"),then_body:vec![expression("branch")],elifs:vec![],else_body:None}),
    ]});
    let collect=|deferred| {let mut names=vec![];super::walk_include_lookahead_exprs(&root,deferred,&mut |e|if let ExprKind::Ident(name)=&e.kind{names.push(name.clone())});names};
    assert_eq!(collect(false),["hot","condition","branch"]);
    let mut reference=vec![];root.walk_exprs(&mut |e|if let ExprKind::Ident(name)=&e.kind{reference.push(name.clone())});
    assert_eq!(collect(true),reference);
}

    use std::path::PathBuf;
    use std::sync::{Arc, Mutex};
    use std::time::{SystemTime, UNIX_EPOCH};

    use super::{JsDynamicRuntime, RuntimeCompilerService, ensure_js_runtime_registered};
    use vybe_runtime::{VM, Value};

    #[test]
    fn completion_lookahead_advances_past_unreached_include_branches() {
        vybe_language_php::register();
        let base=std::env::temp_dir().join(format!("vybe-idle-lookahead-{}",uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        for i in 0..16 {
            std::fs::write(base.join(format!("cold{i}.php")),format!("<?php function unused_cold_{i}() {{ return {i}; }}")).unwrap();
        }
        std::fs::write(base.join("first.php"),"<?php require __DIR__.'/second.php'; function live_first() { return 9; }").unwrap();
        std::fs::write(base.join("second.php"),"<?php function live_second() { return 1; }").unwrap();
        let source=format!("<?php function never_called() {{ {} }} require __DIR__.'/first.php'; return live_first()+live_second();",
            (0..16).map(|i|format!("require __DIR__.'/cold{i}.php';")).collect::<String>());
        let mut vm=configured_vm();
        let (_commands,command_rx)=std::sync::mpsc::channel();
        let (events,_event_rx)=std::sync::mpsc::channel();
        vm.attach_debugger(command_rx,events,false);
        let mut compiled={
            let mut producer=RuntimeCompilerService::new(&mut vm);
            producer.compile_source_by_name(source,"php",base.join("main.php")).unwrap()
        };
        // Ensure current-source lookahead completed before execution. Deferred
        // function bodies must no longer fill this window with cold branches.
        // Handoff remains one-shot; restore its exact ownership after inspection.
        let batch=compiled.prepared_sources.take().unwrap().take();
        assert!(!batch.tickets.is_empty(), "reached include must be prepared");
        assert!(batch.tickets.keys().all(|path| !path.file_name().unwrap().to_string_lossy().starts_with("cold")),
            "deferred branches must not consume the initial preparation window");
        let deadline=std::time::Instant::now()+std::time::Duration::from_secs(10);
        while batch.tickets.values().any(|ticket|!ticket.is_ready()) {
            assert!(std::time::Instant::now()<deadline,"initial speculative jobs must complete");
            std::thread::yield_now();
        }
        compiled.prepared_sources=Some(Arc::new(crate::preparation::PreparedSources::new(
            batch.tickets,batch.plan,batch.deferred,
        ).with_lowerings(batch.lowering_frontend.unwrap(),batch.lowerings)));
        {
            let mut consumer=RuntimeCompilerService::new(&mut vm);
            assert_numeric_value(consumer.run_compiled(compiled).unwrap(),10.0);
        }
        for i in 0..16 {assert!(vm.global(&format!("unused_cold_{i}")).is_none(),"unreached includes must publish no functions");}
        let report=vm.debug_report().unwrap();
        assert!(report.lock().unwrap().completed.iter().any(|(name,_)|name=="parallel include lowering (worker wall)"),
            "retirement must allow a later reached include to use completed lowering");
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn preparation_payload_keeps_lowering_receivers_alive_until_one_shot_consumption() {
        vybe_language_php::register();
        let frontend = crate::languages::find_by_name("php").unwrap();
        let cancelled = Arc::new(std::sync::atomic::AtomicBool::new(false));
        let (_sender, result) = crossbeam_channel::bounded(1);
        let owner = std::collections::HashMap::from([(
            PathBuf::from("handoff.php"),
            super::LoweringTicket { result, cancelled: cancelled.clone(), link_budget: Arc::new(super::AtomicUsize::new(0)) },
        )]);
        let payload = Arc::new(crate::preparation::PreparedSources::new(
            Default::default(), Default::default(), Default::default(),
        ).with_lowerings(frontend, owner));
        let consumer = payload.clone();
        drop(payload);
        assert!(!cancelled.load(std::sync::atomic::Ordering::Relaxed));
        let batch = consumer.take();
        assert_eq!(batch.lowerings.len(), 1);
        assert!(std::ptr::fn_addr_eq(batch.lowering_frontend.unwrap().parse, frontend.parse));
        assert!(consumer.take().lowerings.is_empty(), "only one consumer owns the receivers");
        assert!(!cancelled.load(std::sync::atomic::Ordering::Relaxed));
        drop(batch);
        assert!(cancelled.load(std::sync::atomic::Ordering::Relaxed));
    }

    #[test]
    fn staged_saturated_queue_preserves_ready_ast_without_reparsing() {
        static PARSES: std::sync::atomic::AtomicUsize=std::sync::atomic::AtomicUsize::new(0);
        fn counted_parse(source:&str)->Result<vybe_ast::Module,String> {
            PARSES.fetch_add(1,std::sync::atomic::Ordering::SeqCst);
            let mut module=(crate::languages::find_by_name("php").unwrap().parse)(source)?;
            for statement in &mut module.body {
                if let vybe_ast::StmtKind::ClassDecl{modifiers,..}=&mut statement.kind {
                    modifiers.linkage=vybe_ast::ClassLinkage::Library{identity:"queue-identity".into()};
                }
            }
            Ok(module)
        }
        vybe_language_php::register();
        let mut language=crate::languages::find_by_name("php").unwrap();
        language.parse=counted_parse;
        language.name="staged-queue-php-test";
        vybe_runtime::registry::register_language(language);
        let path=std::env::temp_dir().join(format!("vybe-saturated-{}.php",uuid::Uuid::new_v4()));
        let source="<?php class QueueGuard {} return 7;";
        std::fs::write(&path,source).unwrap();let path=path.canonicalize().unwrap();
        let preparation=crate::preparation::PreparationPool::new(1,1).unwrap();
        let ticket=preparation.submit(path.clone(),language).unwrap();
        let deadline=std::time::Instant::now()+std::time::Duration::from_secs(10);
        while !ticket.is_ready() {assert!(std::time::Instant::now()<deadline);std::thread::yield_now();}
        let lowering=crate::preparation::PreparationPool::new(1,1).unwrap();
        let (started,wait_started)=crossbeam_channel::bounded(1);
        let (release,wait_release)=crossbeam_channel::bounded(1);
        assert!(lowering.submit_lowering(Box::new(move||{started.send(()).unwrap();wait_release.recv().unwrap();})));
        wait_started.recv().unwrap();
        assert!(lowering.submit_lowering(Box::new(||{})));
        let mut vm=configured_vm();
        let mut guard=vybe_runtime::value::Object::new();
        guard.properties.insert(crate::primitives::reflection::FIELD_LIBRARY_IDENTITY.into(),Value::String(Arc::from("queue-identity")));
        vm.set_global("queueguard",Value::Object(vybe_runtime::heap::alloc(guard)));
        let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        runtime.preparing.insert(path.clone(),ticket);
        runtime.stage_ready_lowering_on(&lowering,&vm,language,&path);
        assert!(runtime.lowering.is_empty());assert!(runtime.prepared_ready.contains_key(&path));
        assert!(runtime.prepared_ready[&path].module.as_ref().unwrap().body.iter().any(|s|matches!(&s.kind,
            vybe_ast::StmtKind::ClassDecl{modifiers,..} if matches!(modifiers.linkage,vybe_ast::ClassLinkage::Library{..}))),
            "queue rejection must preserve original library linkage");
        let before=PARSES.load(std::sync::atomic::Ordering::SeqCst);
        let bundle=super::bundle_from_source(source,language,&path);
        let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&path).unwrap();
        assert_eq!(PARSES.load(std::sync::atomic::Ordering::SeqCst),before,"ready AST must survive queue rejection");
        assert_eq!(before,1);
        let mut service=RuntimeCompilerService::new(&mut vm);
        assert_numeric_value(service.run_compiled(compiled).unwrap(),7.0);
        release.send(()).unwrap();std::fs::remove_file(path).unwrap();
    }

    #[test]
    fn staged_harvest_preserves_parent_include_hints_before_draining_ticket() {
        vybe_language_php::register();
        let base=std::env::temp_dir().join(format!("vybe-stage-hints-{}",uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let parent=base.join("parent.php");let child=base.join("child.php");
        std::fs::write(&parent,"<?php require __DIR__ . '/child.php'; return 1;").unwrap();
        std::fs::write(&child,"<?php return 2;").unwrap();
        let parent=parent.canonicalize().unwrap();let child=child.canonicalize().unwrap();
        let language=crate::languages::find_by_name("php").unwrap();
        let pool=crate::preparation::PreparationPool::new(1,4).unwrap();
        let ticket=pool.submit_with_hints(parent.clone(),language,Some(crate::php_include_hints::discover)).unwrap();
        let deadline=std::time::Instant::now()+std::time::Duration::from_secs(10);
        while !ticket.is_ready() {assert!(std::time::Instant::now()<deadline);std::thread::yield_now();}
        let vm=configured_vm();let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        runtime.preparing.insert(parent.clone(),ticket);
        runtime.stage_ready_lowering_on(&pool,&vm,language,&parent);
        assert!(runtime.preparing.contains_key(&child),"harvesting parent AST must schedule its child hint");
        assert!(runtime.preparing.remove(&child).unwrap().wait().unwrap().module.is_ok());
        assert!(vm.global("child").is_none());
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn staged_lowering_revalidates_edits_and_preserves_include_order() {
        vybe_language_php::register();
        let base=std::env::temp_dir().join(format!("vybe-lower-{}",uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let language=crate::languages::find_by_name("php").unwrap();
        let pool=crate::preparation::PreparationPool::new(1,4).unwrap();
        let mut vm=configured_vm();
        let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        let mut paths=Vec::new();
        for (name,source) in [("a.php","<?php return 1;"),("b.php","<?php return 2;")] {
            let path=base.join(name);std::fs::write(&path,source).unwrap();let path=path.canonicalize().unwrap();
            let ticket=pool.submit(path.clone(),language).unwrap();
            let deadline=std::time::Instant::now()+std::time::Duration::from_secs(10);
            while !ticket.is_ready() {assert!(std::time::Instant::now()<deadline);std::thread::yield_now();}
            runtime.preparing.insert(path.clone(),ticket);paths.push(path);
        }
        runtime.stage_ready_lowering_on(&pool,&vm,language,&base);
        assert_eq!(runtime.lowering.len(),2);
        assert!(runtime.php_trait_declarations.is_empty());
        // Consume the second queued file first. Compilation must not execute
        // either file or publish globals merely because its worker finished.
        let bundle=super::bundle_from_source(std::fs::read_to_string(&paths[1]).unwrap(),language,&paths[1]);
        let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&paths[1]).unwrap();
        let mut service=RuntimeCompilerService::new(&mut vm);
        assert_numeric_value(service.run_compiled(compiled).unwrap(),2.0);
        std::fs::write(&paths[0],"<?php return 9;").unwrap();
        let bundle=super::bundle_from_source(std::fs::read_to_string(&paths[0]).unwrap(),language,&paths[0]);
        let compiled=runtime.compile_dynamic_php(&mut vm,&bundle,&paths[0]).unwrap();
        let mut service=RuntimeCompilerService::new(&mut vm);
        assert_numeric_value(service.run_compiled(compiled).unwrap(),9.0);
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn staged_lowering_refuses_traits_and_unknown_source_dependencies() {
        vybe_language_php::register();
        let language=crate::languages::find_by_name("php").unwrap();
        for source in ["<?php trait T { function f() {return 1;} }", "<?php class C {use Missing;}","<?php if(false) {trait T {}}"] {
            let bundle=super::bundle_from_source(source,language,PathBuf::from("declined.php"));
            let module=bundle.prepared_module().unwrap();assert!(!super::independent_lowering(&module));
        }
        let base=std::env::temp_dir().join(format!("vybe-lower-deps-{}",uuid::Uuid::new_v4()));
        std::fs::write(&base,"<?php return 7;").unwrap();let path=base.canonicalize().unwrap();
        let pool=crate::preparation::PreparationPool::new(1,1).unwrap();
        let mut prepared=pool.submit(path.clone(),language).unwrap().wait().unwrap();
        prepared.dependencies=Some(vec![PathBuf::from("changing-dependency.php")]);
        // This same source validation predicate gates admission and reuse.
        assert!(!prepared.matches_source(&path,"<?php return 7;"));
        std::fs::remove_file(base).unwrap();
    }

    #[test]
    fn staged_lowering_never_publishes_an_unreached_parse_failure() {
        vybe_language_php::register();
        let base=std::env::temp_dir().join(format!("vybe-lower-invalid-{}",uuid::Uuid::new_v4()));
        std::fs::write(&base,"<?php function (").unwrap();let path=base.canonicalize().unwrap();
        let language=crate::languages::find_by_name("php").unwrap();
        let pool=crate::preparation::PreparationPool::new(1,1).unwrap();
        let ticket=pool.submit(path.clone(),language).unwrap();
        let deadline=std::time::Instant::now()+std::time::Duration::from_secs(10);
        while !ticket.is_ready() {assert!(std::time::Instant::now()<deadline);std::thread::yield_now();}
        let vm=configured_vm();let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        runtime.preparing.insert(path.clone(),ticket);runtime.stage_ready_lowering_on(&pool,&vm,language,&base);
        assert!(runtime.lowering.is_empty());assert!(runtime.prepared_ready[&path].module.is_err());
        assert!(runtime.php_trait_declarations.is_empty());assert!(runtime.once_resources.is_empty());
        std::fs::remove_file(base).unwrap();
    }

    #[test]
    fn staged_context_checks_negative_installed_library_decisions() {
        vybe_language_php::register();
        let language=crate::languages::find_by_name("php").unwrap();
        let mut vm=VM::new();
        let context=super::CompilationDependencies {
            certificate:Arc::new(Mutex::new(crate::compile_dependencies::Certificate::default())),
            libraries:vec![("LibraryGuard".into(),"identity-v1".into(),false)],profile:(language.profile_source)(),profile_source:language.profile_source,language_name:language.name,frontend:super::FrontendIdentity::of(language),
        };
        assert!(context.valid(&vm));
        let mut object=vybe_runtime::value::Object::new();
        object.properties.insert(crate::primitives::reflection::FIELD_LIBRARY_IDENTITY.into(),Value::String(Arc::from("identity-v1")));
        vm.set_global("libraryguard",Value::Object(vybe_runtime::heap::alloc(object)));
        assert!(!context.valid(&vm));
    }

    #[test]
    fn speculative_compiler_diagnostics_stay_private_until_consumed() {
        let certificate=Arc::new(Mutex::new(crate::compile_dependencies::Certificate::speculative()));
        assert!(certificate.lock().unwrap().capture_diagnostic("deferred diagnostic"));
        assert_eq!(certificate.lock().unwrap().take_diagnostics(),["deferred diagnostic"]);
        assert!(certificate.lock().unwrap().take_diagnostics().is_empty());
    }

    struct DynamicSmokeCase {
        language: &'static str,
        virtual_path: &'static str,
        source: &'static str,
    }

    fn configured_vm() -> VM {
        let mut vm = VM::new();
        crate::primitives::platforms::register_platforms_all(&mut vm);
        vm
    }

    fn assert_numeric_value(value: Value, expected: f64) {
        match value {
            Value::I32(n) => assert_eq!(n as f64, expected),
            Value::I64(n) => assert_eq!(n as f64, expected),
            Value::F64(n) => assert_eq!(n, expected),
            other => panic!("expected numeric value, got {other:?}"),
        }
    }

    #[test]
    fn library_classes_reuse_installed_code_and_keep_inherited_members() {
        let mut vm = configured_vm();
        let language = crate::languages::find_by_name("php").unwrap();
        let mut service = RuntimeCompilerService::new(&mut vm);
        for (base, first_source, second_source, expected) in [
            ("runtimeexception",
             "<?php return (new RuntimeException('first'))->getMessage();",
             "<?php class ReusedException extends RuntimeException {} return (new ReusedException('second'))->getMessage();",
             "second"),
            ("splfileinfo",
             "<?php return (new SplFileInfo('first.php'))->getFilename();",
             "<?php class ReusedFile extends SplFileInfo {} return (new ReusedFile('second.php'))->getFilename();",
             "second.php"),
        ] {
            let first = service.compile_source(first_source, language, "library-first.php").unwrap();
            let first_chunks = first.chunks.len();
            let first_value = service.run_compiled(first).unwrap();
            assert!(matches!(first_value, Value::String(_)));
            let installed = service.vm().global(base).unwrap();
            assert!(super::object_get_prop(installed,
                crate::primitives::reflection::FIELD_LIBRARY_IDENTITY).is_some());
            let second = service.compile_source(second_source, language, "library-second.php").unwrap();
            assert!(second.chunks.len() < first_chunks,
                "installed library methods must not be emitted again");
            let second_value = service.run_compiled(second).unwrap();
            assert!(matches!(second_value, Value::String(ref value) if value.as_ref() == expected),
                "inherited library method returned {second_value:?}");
        }
    }

    #[test]
    fn compiles_source_into_live_vm_and_publishes_function_globals() {
        let mut vm = configured_vm();

        {
            let mut service = RuntimeCompilerService::new(&mut vm);
            service
                .compile_and_run_source(
                    "function greet() { return 7; }",
                    crate::languages::find_by_name("js").expect("js language"),
                    PathBuf::from("dynamic/greet.js"),
                )
                .expect("compile and run greet");

            service
                .compile_and_run_source(
                    "function callGreet() { return greet(); }",
                    crate::languages::find_by_name("js").expect("js language"),
                    PathBuf::from("dynamic/call_greet.js"),
                )
                .expect("compile and run callGreet");
        }

        let greet = vm.global("greet").cloned().expect("greet global");
        let call_greet = vm.global("callgreet").cloned().expect("callGreet global");

        assert_numeric_value(vm.invoke(&greet, &[]).expect("invoke greet"), 7.0);
        assert_numeric_value(vm.invoke(&call_greet, &[]).expect("invoke callGreet"), 7.0);
    }

    #[test]
    fn dynamic_execution_smoke_matrix_for_supported_languages() {
        let cases = [
            DynamicSmokeCase {
                language: "js",
                virtual_path: "dynamic/matrix.js",
                source: "let x = 7;",
            },
            DynamicSmokeCase {
                language: "php",
                virtual_path: "dynamic/matrix.php",
                source: "<?php $x = 7;",
            },
            DynamicSmokeCase {
                language: "python",
                virtual_path: "dynamic/matrix.py",
                source: "x = 7",
            },
            DynamicSmokeCase {
                language: "ruby",
                virtual_path: "dynamic/matrix.rb",
                source: "x = 7",
            },
            DynamicSmokeCase {
                language: "dart",
                virtual_path: "dynamic/matrix.dart",
                source: "var x = 7;",
            },
            DynamicSmokeCase {
                language: "vb",
                virtual_path: "dynamic/matrix.vb",
                source: "Dim x As Integer = 7",
            },
            DynamicSmokeCase {
                language: "csharp",
                virtual_path: "dynamic/matrix.cs",
                source: "int x = 7;",
            },
            DynamicSmokeCase {
                language: "pascal",
                virtual_path: "dynamic/matrix.pas",
                source: "program T; var x: Integer; begin x := 7; end.",
            },
            DynamicSmokeCase {
                language: "cobol",
                virtual_path: "dynamic/matrix.cob",
                source: "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC 9 VALUE 7.\nPROCEDURE DIVISION.\n    STOP RUN.",
            },
            DynamicSmokeCase {
                language: "fortran",
                virtual_path: "dynamic/matrix.f90",
                source: "program test\n  integer :: x\n  x = 7\nend program test",
            },
        ];

        for case in cases {
            let mut vm = configured_vm();
            let mut service = RuntimeCompilerService::new(&mut vm);
            let compiled = service
                .compile_source_by_name(
                    case.source,
                    case.language,
                    PathBuf::from(case.virtual_path),
                )
                .unwrap_or_else(|err| panic!("{} dynamic compile failed: {}", case.language, err));

            service
                .run_compiled(compiled)
                .unwrap_or_else(|err| panic!("{} dynamic run failed: {}", case.language, err));
        }
    }

    #[test]
    fn dynamic_compile_requires_capability_for_source_text() {
        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::with_capabilities(
            &mut vm,
            vybe_runtime::capabilities::Capabilities::safe(),
        );

        let err = service
            .compile_source_by_name("let x = 7;", "js", PathBuf::from("dynamic/locked.js"))
            .expect_err("dynamic compile should be denied without capability");

        assert!(err.contains("DynamicCompile"), "unexpected error: {err}");
    }

    #[test]
    fn js_function_constructor_bridge_returns_callable_function() {
        let mut vm = configured_vm();
        ensure_js_runtime_registered(&mut vm);

        let value = {
            let mut runtime =
                JsDynamicRuntime::new(vybe_runtime::capabilities::Capabilities::all());
            let _guard = runtime.activate(&mut vm);
            runtime
                .handle_function_constructor(&[
                    Value::String("a".into()),
                    Value::String("b".into()),
                    Value::String("return a + b;".into()),
                ])
                .expect("construct function")
        };

        match vm.invoke(&value, &[Value::I32(2), Value::I32(3)]) {
            Ok(Value::I32(value)) => assert_eq!(value, 5),
            Ok(Value::I64(value)) => assert_eq!(value, 5),
            Ok(Value::F64(value)) => assert_eq!(value, 5.0),
            Ok(other) => panic!("expected numeric result, got {other:?}"),
            Err(err) => panic!("expected callable dynamic function, got {err}"),
        }
    }

    #[test]
    fn js_function_constructor_call_import_returns_callable_function() {
        let mut vm = configured_vm();
        ensure_js_runtime_registered(&mut vm);
        let host_idx = *vm
            .host_registry
            .get(&("vybe:js".to_string(), "function_constructor".to_string()))
            .expect("vybe:js:function_constructor host fn");

        let mut chunk = vybe_runtime::chunk::Chunk::new("<script>");
        let import_idx = chunk.add_import("vybe:js", "function_constructor");
        // String constants go through `wasm:string-constants`, the same route
        // the compiler's `push_const` uses — not the custom `CONST` opcode.
        chunk.emit_string_const("a", 0);
        chunk.emit_string_const("b", 0);
        chunk.emit_string_const("return a + b;", 0);
        chunk.emit_call(import_idx, 3, 0);
        crate::primitives::globals::emit_write(&mut chunk, "__test_result", 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);

        {
            let mut runtime =
                JsDynamicRuntime::new(vybe_runtime::capabilities::Capabilities::all());
            let _guard = runtime.activate(&mut vm);
            vm.run_linked(
                vec![chunk],
                vec![vybe_runtime::ImportTarget::Host(host_idx)],
            )
            .expect("call vybe:js:function_constructor")
        };

        let value = vm
            .global("__test_result")
            .cloned()
            .expect("stored host-import result");

        match &value {
            Value::Object(_) => {}
            other => panic!("expected object result from CALL_IMPORT, got {other:?}"),
        }

        match vm.invoke(&value, &[Value::I32(2), Value::I32(3)]) {
            Ok(Value::I32(value)) => assert_eq!(value, 5),
            Ok(Value::I64(value)) => assert_eq!(value, 5),
            Ok(Value::F64(value)) => assert_eq!(value, 5.0),
            Ok(other) => panic!("expected numeric result, got {other:?}"),
            Err(err) => panic!("expected callable import result, got {err}"),
        }
    }

    #[test]
    fn php_dynamic_include_executes_expression_form() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-include-{stamp}"));
        std::fs::create_dir_all(&base).expect("create temp dir");

        let main_path = base.join("main.php");
        let partial_path = base.join("partial.php");

        std::fs::write(
            &main_path,
            "<?php $path = 'partial.php'; $value = include $path; $result = get_value() + $value; return $result;",
        )
        .expect("write main php");
        std::fs::write(
            &partial_path,
            "<?php function get_value() { return 8; } return 34;",
        )
        .expect("write partial php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        let result = service
            .compile_and_run_path(&main_path)
            .expect("run php with dynamic include");

        assert_numeric_value(result, 42.0);
    }

    #[test]
    fn php_worker_discovery_schedules_descendants_without_publishing_or_raising_errors() {
        vybe_language_php::register();
        let base = std::env::temp_dir().join(format!("vybe-worker-hints-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let main = base.join("main.php");
        let parent = base.join("parent.php");
        let child = base.join("child.php");
        let invalid = base.join("invalid.php");
        std::fs::write(&parent, "<?php require ROOT . '/child.php'; if (false) require __DIR__ . '/invalid.php';").unwrap();
        std::fs::write(&child, "<?php function must_not_publish() { return 7; }").unwrap();
        std::fs::write(&invalid, "<?php function (").unwrap();
        let parent = parent.canonicalize().unwrap();
        let language = crate::languages::find_by_name("php").unwrap();
        let pool = crate::preparation::PreparationPool::new(1, 1).unwrap();
        let ticket = pool.submit_with_hints(parent.clone(), language,
            Some(crate::php_include_hints::discover)).unwrap();
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(10);
        while !ticket.is_ready() && std::time::Instant::now() < deadline {
            std::thread::sleep(std::time::Duration::from_millis(1));
        }
        assert!(ticket.is_ready());
        let mut runtime = super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        runtime.preparing.insert(parent.clone(), ticket);
        let mut vm = VM::new();
        vm.set_global("ROOT", Value::String(Arc::from(base.to_string_lossy().as_ref())));
        runtime.refresh_worker_preparation_hints(&main, language, &vm);
        assert!(runtime.preparing.remove(&child.canonicalize().unwrap()).unwrap().wait().unwrap().module.is_ok());
        assert!(runtime.preparing.remove(&invalid.canonicalize().unwrap()).unwrap().wait().unwrap().module.is_err(),
            "an unreached syntax error remains a speculative result");
        assert!(vm.global("must_not_publish").is_none(), "preparation never installs declarations");
        assert!(runtime.preparing.remove(&parent).unwrap().wait().unwrap().module.is_ok());
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn php_include_preparation_starts_before_parent_ast_preparation() {
        vybe_language_php::register();
        let base = std::env::temp_dir().join(format!("vybe-prepare-early-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let main = base.join("main.php");
        let child = base.join("child.php");
        std::fs::write(&child, "<?php return 42;").unwrap();
        // No parent parse or Module exists at this point. Invalid trailing
        // syntax must not prevent discovery of an earlier literal include.
        let bundle = super::bundle_from_source("<?php require __DIR__ . '/child.php'; function (",
            crate::languages::find_by_name("php").unwrap(), &main);
        let mut runtime = super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        runtime.prepare_php_source_hints(&bundle, &main, None);
        let ticket = runtime.preparing.remove(&child.canonicalize().unwrap())
            .expect("child preparation must start without a parent AST");
        let prepared = ticket.wait().expect("child preparation completes");
        assert!(prepared.module.is_ok());
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn php_preparation_retries_parent_hints_after_constants_become_available() {
        vybe_language_php::register();
        let base = std::env::temp_dir().join(format!("vybe-prepare-late-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let main = base.join("settings.php");
        let child = base.join("child.php");
        std::fs::write(&main, "<?php require ROOT . '/child.php';").unwrap();
        std::fs::write(&child, "<?php return 42;").unwrap();
        let main = main.canonicalize().unwrap();
        let child = child.canonicalize().unwrap();
        let language = crate::languages::find_by_name("php").unwrap();
        let bundle = super::bundle_from_source(std::fs::read_to_string(&main).unwrap(), language, &main);
        let mut runtime = super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        let mut vm = VM::new();
        runtime.prepare_php_source_hints(&bundle, &main, Some(&vm));
        assert!(runtime.preparing.is_empty());
        vm.set_global("ROOT", Value::String(Arc::from(base.to_string_lossy().as_ref())));
        runtime.refresh_php_preparation_hints(&main, &main, language, &vm);
        let ticket = runtime.preparing.remove(&child).expect("late constant schedules child");
        let prepared = ticket.wait().expect("worker prepares child");
        assert!(prepared.module.is_ok());
        runtime.refresh_php_preparation_hints(&main, &main, language, &vm);
        assert!(runtime.preparing.is_empty(), "resolved hint is consumed once");
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn php_preparation_survives_compiler_service_handoff() {
        let base = std::env::temp_dir().join(format!("vybe-prepare-handoff-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let main = base.join("main.php");
        std::fs::write(&main, "<?php return include 'child.php';").unwrap();
        std::fs::write(base.join("child.php"), "<?php return 42;").unwrap();
        let mut vm = configured_vm();
        let compiled = {
            let mut service = RuntimeCompilerService::new(&mut vm);
            service.compile_path(&main).unwrap()
        };
        assert!(compiled.prepared_sources.is_some(), "entry must carry its queued preparation");
        let mut service = RuntimeCompilerService::new(&mut vm);
        let value = service.run_compiled(compiled).unwrap();
        assert_numeric_value(value, 42.0);
        assert!(service.php_runtime.preparing.is_empty(), "included child must consume its preparation");
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn php_parallel_preparation_rechecks_edits_and_defers_unreached_errors() {
        let base = std::env::temp_dir().join(format!("vybe-prepare-php-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let main = base.join("main.php");
        let child = base.join("child.php");
        std::fs::write(&main,
            "<?php if (false) { include 'broken.php'; } return include 'child.php';").unwrap();
        std::fs::write(base.join("broken.php"), "<?php function (").unwrap();
        std::fs::write(&child, "<?php return 1;").unwrap();
        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        let compiled = service.compile_path(&main).expect("unreached include must not fail entry compilation");
        // Use an owned queue so other concurrent tests cannot decline this job.
        let pool = crate::preparation::PreparationPool::new(1, 2).unwrap();
        let ticket = pool.submit(child.clone(), crate::languages::find_by_name("php").unwrap()).unwrap();
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(10);
        while !ticket.is_ready() && std::time::Instant::now() < deadline {
            std::thread::sleep(std::time::Duration::from_millis(1));
        }
        assert!(ticket.is_ready(), "prepare original child before editing it");
        service.php_runtime.preparing.insert(child.canonicalize().unwrap(), ticket);
        std::fs::write(&child, "<?php return 42;").unwrap();
        let value = service.run_compiled(compiled).expect("edited child must prepare again; unreached error stays deferred");
        assert_numeric_value(value, 42.0);
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn cached_trait_unit_restores_declarations_for_a_fresh_consumer() {
        #[derive(Default)]
        struct Cache {
            units: Mutex<std::collections::HashMap<std::path::PathBuf, super::DynamicCompilation>>,
            hits: std::sync::atomic::AtomicUsize,
        }
        impl super::IncludeCompileCache for Cache {
            fn get(&self, path: &std::path::Path, _: u64) -> Option<Result<super::DynamicCompilation, String>> {
                let unit = self.units.lock().unwrap().get(path).cloned()?;
                self.hits.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
                Some(Ok(unit))
            }
            fn store(&self, path: &std::path::Path, _: u64, _: Option<Vec<std::path::PathBuf>>,
                result: Result<&super::DynamicCompilation, &str>) {
                if let Ok(unit) = result {
                    // Deliberately recompile the consumer each request. Replaying
                    // its already-augmented chunks would conceal the missing map.
                    if !unit.source_declarations.is_empty() {
                        self.units.lock().unwrap().insert(path.into(), unit.clone());
                    }
                }
            }
        }
        let base = std::env::temp_dir().join(format!("vybe-trait-cache-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let main = base.join("main.php");
        let provider = base.join("provider.php");
        std::fs::write(&provider, "<?php trait CachedSource { public function value() { return 42; } }").unwrap();
        std::fs::write(base.join("consumer.php"), "<?php class FreshConsumer { use CachedSource; } return (new FreshConsumer())->value();").unwrap();
        std::fs::write(&main, "<?php require_once 'provider.php'; return include 'consumer.php';").unwrap();
        let cache = Arc::new(Cache::default());
        let mut declarations = Vec::new();
        for _ in 0..2 {
            let mut vm = configured_vm();
            let mut service = RuntimeCompilerService::new(&mut vm);
            service.set_include_cache(cache.clone());
            assert_numeric_value(service.compile_and_run_path(&main).unwrap(), 42.0);
            let declaration = service.php_runtime.php_trait_declarations.get("cachedsource")
                .expect("cached provider must restore the same compile-time schema");
            declarations.push(format!("{declaration:?}"));
        }
        assert_eq!(declarations[0], declarations[1], "cold/warm source AST is identical");
        assert_eq!(cache.hits.load(std::sync::atomic::Ordering::Relaxed), 1);
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn compiled_trait_preparation_does_not_publish_request_metadata() {
        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        let language = crate::languages::find_by_name("php").unwrap();
        let unit = service.compile_source("<?php trait DeferredSource { public function value() { return 7; } }",
            language, "deferred-source.php").unwrap();
        assert_eq!(unit.source_declarations.len(), 1);
        assert!(service.php_runtime.php_trait_declarations.is_empty());
        assert!(service.vm.global("DeferredSource").is_none());
        service.run_compiled(unit).unwrap();
        assert!(service.php_runtime.php_trait_declarations.contains_key("deferredsource"));
    }

    #[test]
    fn cached_library_references_recompile_when_the_new_vm_lacks_the_provider() {
        struct Cache {
            compiled: Mutex<super::DynamicCompilation>,
            stores: std::sync::atomic::AtomicUsize,
        }
        impl super::IncludeCompileCache for Cache {
            fn get(&self, _: &std::path::Path, _: u64) -> Option<Result<super::DynamicCompilation, String>> {
                Some(Ok(self.compiled.lock().unwrap().clone()))
            }
            fn store(&self, _: &std::path::Path, _: u64, _: Option<Vec<PathBuf>>,
                result: Result<&super::DynamicCompilation, &str>) {
                *self.compiled.lock().unwrap() = result.unwrap().clone();
                self.stores.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            }
        }
        let base = std::env::temp_dir().join(format!("vybe-library-cache-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let part = base.join("part.php");
        let main = base.join("main.php");
        let source = "<?php return (new SplFileInfo('cached.php'))->getFilename();";
        std::fs::write(&part, source).unwrap();
        std::fs::write(&main, "<?php $path = 'part.php'; return include $path;").unwrap();
        let mut warm_vm = configured_vm();
        let language = crate::languages::find_by_name("php").unwrap();
        let mut warm = RuntimeCompilerService::new(&mut warm_vm);
        warm.compile_and_run_source(source, language, &part).unwrap();
        let cached = warm.compile_source(source, language, &part).unwrap();
        assert!(cached.required_libraries.iter().any(|library| library.name == "SplFileInfo"));
        let mut cold_vm = configured_vm();
        assert!(!cached.libraries_available(&cold_vm));
        let cache = Arc::new(Cache { compiled: Mutex::new(cached),
            stores: std::sync::atomic::AtomicUsize::new(0) });
        let mut cold = RuntimeCompilerService::new(&mut cold_vm);
        cold.set_include_cache(cache.clone());
        let result = cold.compile_and_run_path(&main).unwrap();
        assert!(matches!(result, Value::String(ref value) if value.as_ref() == "cached.php"));
        assert_eq!(cache.stores.load(std::sync::atomic::Ordering::Relaxed), 1,
            "incompatible cached references must compile again with their provider");
        assert!(!cache.compiled.lock().unwrap().required_libraries.iter()
            .any(|library| library.name == "SplFileInfo"));
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn php_cached_plain_includes_execute_again_and_once_state_resets_per_run() {
        #[derive(Default)]
        struct Cache {
            compiled: Mutex<Option<super::DynamicCompilation>>,
            hits: std::sync::atomic::AtomicUsize,
            stores: std::sync::atomic::AtomicUsize,
        }
        impl super::IncludeCompileCache for Cache {
            fn get(&self, _: &std::path::Path, _: u64)
                -> Option<Result<super::DynamicCompilation, String>>
            {
                let compiled = self.compiled.lock().unwrap().clone()?;
                self.hits.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
                Some(Ok(compiled))
            }
            fn store(&self, _: &std::path::Path, _: u64,
                deps: Option<Vec<PathBuf>>,
                result: Result<&super::DynamicCompilation, &str>)
            {
                assert!(deps.is_some());
                *self.compiled.lock().unwrap() = Some(result.unwrap().clone());
                self.stores.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            }
        }
        let base = std::env::temp_dir().join(format!("vybe-include-cache-{}", uuid::Uuid::new_v4()));
        std::fs::create_dir_all(&base).unwrap();
        let main = base.join("main.php");
        std::fs::write(&main,
            "<?php $a = include 'part.php'; $b = include 'part.php'; $c = include_once 'part.php'; return $a + $b + $c;").unwrap();
        std::fs::write(base.join("part.php"), "<?php return 21;").unwrap();
        let cache = Arc::new(Cache::default());
        for _ in 0..2 {
            let mut vm = configured_vm();
            let mut service = RuntimeCompilerService::new(&mut vm);
            service.set_include_cache(cache.clone());
            let value = service.compile_and_run_path(&main).expect("run cached includes");
            assert_numeric_value(value, 43.0);
        }
        assert_eq!(cache.stores.load(std::sync::atomic::Ordering::Relaxed), 1);
        assert_eq!(cache.hits.load(std::sync::atomic::Ordering::Relaxed), 3);
        std::fs::remove_dir_all(base).unwrap();
    }

    #[test]
    fn php_includes_reuse_installed_exception_bootstrap() {
        let stamp = SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-php-shared-exceptions-{stamp}"));
        std::fs::create_dir_all(&base).unwrap();
        let main = base.join("main.php");
        std::fs::write(&main, "<?php $a = 'a.php'; include $a; $b = 'b.php'; include $b;",).unwrap();
        std::fs::write(base.join("a.php"), "<?php $first = new Exception('one');").unwrap();
        std::fs::write(base.join("b.php"), "<?php $second = new Exception('two');").unwrap();

        let mut vm = configured_vm();
        RuntimeCompilerService::new(&mut vm).compile_and_run_path(&main).unwrap();
        let exception_chunks = vm.chunks.iter().filter(|chunk| chunk.name == "Exception").count();
        assert_eq!(exception_chunks, 1, "core Exception must be compiled once per VM");
    }

    #[test]
    fn php_dynamic_include_resolves_magic_file_constants_in_nested_runtime_paths() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-magic-include-{stamp}"));
        std::fs::create_dir_all(&base).expect("create temp dir");

        let main_path = base.join("main.php");
        let child_path = base.join("child.php");
        let grandchild_path = base.join("grandchild.php");

        std::fs::write(
            &main_path,
            "<?php $path = 'child.php'; $result = include $path; return $result;",
        )
        .expect("write main php");
        std::fs::write(
            &child_path,
            "<?php $name = 'grandchild'; return include __DIR__ . '/' . $name . '.php';",
        )
        .expect("write child php");
        std::fs::write(
            &grandchild_path,
            "<?php return strpos(__FILE__, 'grandchild.php') !== false ? 42 : 0;",
        )
        .expect("write grandchild php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        let result = service
            .compile_and_run_path(&main_path)
            .expect("run nested php dynamic include");

        assert_numeric_value(result, 42.0);
    }

    #[test]
    fn php_dynamic_plain_include_active_cycle_returns_success_without_reentering() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-active-include-{stamp}"));
        std::fs::create_dir_all(&base).expect("create temp dir");

        let main_path = base.join("main.php");
        let child_path = base.join("child.php");

        std::fs::write(
            &main_path,
            "<?php $first = require __DIR__ . '/child.php'; $second = require __DIR__ . '/child.php'; $result = $first + $second; return $result;",
        )
        .expect("write main php");
        std::fs::write(
            &child_path,
            "<?php $nested = require __FILE__; return $nested + 40;",
        )
        .expect("write child php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        let result = service
            .compile_and_run_path(&main_path)
            .expect("run php active include cycle");

        assert_numeric_value(result, 82.0);
    }

    #[test]
    fn php_dynamic_include_normalizes_alternative_control_syntax() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-alt-syntax-{stamp}"));
        std::fs::create_dir_all(&base).expect("create temp dir");

        let main_path = base.join("main.php");
        let compat_path = base.join("compat.php");

        std::fs::write(
            &main_path,
            "<?php $result = require __DIR__ . '/compat.php'; $call_result = compat_polyfill(); return ($result === 42 && $call_result === 42) ? 42 : 0;",
        )
        .expect("write main php");
        std::fs::write(
            &compat_path,
            "<?php\nif ( ! function_exists( 'compat_polyfill' ) ) :\n\t/**\n\t * Timing attack safe string comparison.\n\t *\n\t * @param string $known_string Expected string.\n\t * @param string $user_string  Actual, user supplied, string.\n\t * @return bool Whether strings are equal.\n\t */\n\tfunction compat_polyfill( $known_string = 'a', $user_string = 'a' ) {\n\t\t$known_string_length = strlen( $known_string );\n\n\t\tif ( strlen( $user_string ) !== $known_string_length ) {\n\t\t\treturn false;\n\t\t}\n\n\t\t$result = 0;\n\n\t\tfor ( $i = 0; $i < $known_string_length; $i++ ) {\n\t\t\t$result |= ord( $known_string[ $i ] ) ^ ord( $user_string[ $i ] );\n\t\t}\n\n\t\treturn 0 === $result ? 42 : 0;\n\t}\nendif;\n\nreturn compat_polyfill();\n",
        )
        .expect("write compat php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        let result = service
            .compile_and_run_path(&main_path)
            .expect("run php with alternative syntax include");

        assert_numeric_value(result, 42.0);

    }

    #[test]
    fn php_dynamic_include_uses_entry_script_dir_for_nested_bare_relative_paths() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-entry-relative-{stamp}"));
        let classes_dir = base.join("classes");
        let views_dir = base.join("views");
        std::fs::create_dir_all(&classes_dir).expect("create classes dir");
        std::fs::create_dir_all(&views_dir).expect("create views dir");

        let main_path = base.join("main.php");
        let child_path = classes_dir.join("child.php");
        let partial_path = views_dir.join("partial.php");

        std::fs::write(
            &main_path,
            "<?php $path = 'classes/child.php'; $result = include $path; return $result;",
        )
        .expect("write main php");
        std::fs::write(
            &child_path,
            "<?php $view = 'views/partial.php'; return include $view;",
        )
        .expect("write child php");
        std::fs::write(&partial_path, "<?php return 42;").expect("write partial php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        let result = service
            .compile_and_run_path(&main_path)
            .expect("run php with entry-relative nested include");

        assert_numeric_value(result, 42.0);
    }

    #[test]
    fn php_dynamic_include_renders_nested_view_templates_with_entry_relative_paths() {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let base = std::env::temp_dir().join(format!("vybe-dynamic-view-render-{stamp}"));
        let views_projects_dir = base.join("views/projects");
        let views_templates_dir = base.join("views/templates");
        std::fs::create_dir_all(&views_projects_dir).expect("create projects view dir");
        std::fs::create_dir_all(&views_templates_dir).expect("create templates dir");

        let main_path = base.join("index.php");
        let list_path = views_projects_dir.join("list.php");
        let header_path = views_templates_dir.join("header.php");

        std::fs::write(
            &main_path,
            "<?php $result = include 'views/projects/list.php';",
        )
        .expect("write main php");
        std::fs::write(
            &list_path,
            "<?php $appName = 'Plan'; include 'views/templates/header.php'; ?>\n<main>Projects</main>\n",
        )
        .expect("write list php");
        std::fs::write(
            &header_path,
            "<header><?= htmlspecialchars($appName) ?></header>\n",
        )
        .expect("write header php");

        let rendered = Arc::new(Mutex::new(String::new()));
        let rendered_sink = rendered.clone();
        let stdout_sink = rendered.clone();

        let mut vm = configured_vm();
        // PHP echo writes through the WASI stdout stream. Capture that stream
        // in addition to console calls so these assertions inspect the real
        // rendered include output rather than an unused host interface.
        vm.register_host_fn(
            "wasi:cli/stdout",
            "write-via-stream",
            Box::new(move |ctx, args| {
                let stream = args.first().cloned().unwrap_or(Value::Null);
                let bytes = ctx.stream_drain(&stream);
                stdout_sink.lock().expect("lock stdout output")
                    .push_str(&String::from_utf8_lossy(&bytes));
                let (future, id) = ctx.create_future();
                ctx.resolve_future(id, Value::Null);
                future
            }),
        );
        vm.register_host_fn(
            "web:console",
            "log",
            Box::new(move |_ctx, args| {
                let mut out = rendered_sink.lock().expect("lock rendered output");
                for arg in args {
                    match arg {
                        Value::String(text) => out.push_str(text.as_ref()),
                        other => out.push_str(&format!("{}", other)),
                    }
                }
                Value::Null
            }),
        );

        let mut service = RuntimeCompilerService::new(&mut vm);
        service
            .compile_and_run_path(&main_path)
            .expect("run php nested view include");

        let rendered = rendered.lock().expect("lock rendered output").clone();
        assert!(
            rendered.contains("<header>Plan</header>"),
            "missing rendered header: {rendered}"
        );
        assert!(
            rendered.contains("<main>Projects</main>"),
            "missing rendered list body: {rendered}"
        );
    }
    #[test]
    fn immediate_and_certified_include_chunks_errors_match() {
        vybe_language_php::register();
        let _cap = crate::primitives::instructions::host::CapabilityContext::get();
        let language=crate::languages::find_by_name("php").unwrap();
        let cases=[
            "<?php return 7;",
            "<?php function f($a=2) { return $a+3; } return f();",
            "<?php class A { public $x=4; function f() {return $this->x;} } return (new A())->f();",
            "<?php trait T { function f(){return 8;} } class A {use T;} return (new A())->f();",
            "<?php namespace N; function f($x){return $x+1;} return f(4);",
            "<?php function (",
            "<?php class A {use MissingTrait;} return 1;",
        ];
        for (n,source) in cases.iter().enumerate() {
            let path=std::env::temp_dir().join(format!("vybe-immediate-private-{n}.php"));
            let bundle=super::bundle_from_source(*source,language,path.clone());
            let mut a=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
            let mut b=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
            let mut avm=configured_vm();let mut bvm=configured_vm();
            let left=a.compile_dynamic_php_with_reuse(&mut avm,&bundle,&path,true);
            let right=b.compile_dynamic_php_with_reuse(&mut bvm,&bundle,&path,false);
            match (left,right) {
                (Ok(x),Ok(y))=>{
                    assert!(x.compile_dependencies.is_some());assert!(y.compile_dependencies.is_none());
                    assert_eq!(x.chunks.len(),y.chunks.len(),"case {n}");
                    for (cx,cy) in x.chunks.iter().zip(&y.chunks) {
                        assert_eq!(cx.code,cy.code,"case {n} chunk {}",cx.name);
                        assert_eq!(format!("{:?}",(&cx.constants,&cx.imports,&cx.global_imports,&cx.globals,&cx.types,&cx.lines,&cx.local_names,&cx.canon_section)),format!("{:?}",(&cy.constants,&cy.imports,&cy.global_imports,&cy.globals,&cy.types,&cy.lines,&cy.local_names,&cy.canon_section)),"case {n}");
                    }
                    assert_eq!(format!("{:?}",(&x.source_declarations,&x.required_libraries,&x.host_imports,&x.component_binary,&x.app_shell)),format!("{:?}",(&y.source_declarations,&y.required_libraries,&y.host_imports,&y.component_binary,&y.app_shell)),"case {n}");
                },
                (Err(x),Err(y))=>assert_eq!(x,y,"case {n}"),
                _=>panic!("include outcome mismatch case {n}"),
            }
            assert_eq!(a.contextual_compile_failure,b.contextual_compile_failure,"case {n}");
        }
    }

    #[test]
    fn prepared_ast_emission_preserves_source_bundle_code_and_metadata() {
        vybe_language_php::register();
        let _caps = crate::primitives::instructions::host::CapabilityContext::get();
        let language = crate::languages::find_by_name("php").unwrap();
        let vm = configured_vm();
        let cases = [
            "<?php return 7;",
            "<?php function f($a=2) { return $a+3; } return f();",
            "<?php namespace N; class A { public $x=4; function f(){return $this->x;} } return (new A())->f();",
            "<header>before</header><?php echo 'inside'; ?><footer>after</footer>",
            "<?php function f($path){$x=1; include $path; include_once $path; require $path; require_once $path; return $x;}",
            "<?php $value=null; return 'prefix'.$value;",
        ];
        for (index, source) in cases.iter().enumerate() {
            let path = std::env::temp_dir().join(format!("vybe-prepared-ast-emission-{index}.php"));
            let module = (language.parse)(source).expect("parse real PHP fixture");
            let bundle = super::bundle_from_source(*source, language, &path);
            let control = bundle.compile_prepared_certified_with_helpers(
                &module, &vm.modules, false, None, None, None).expect("emit source bundle");
            let direct = crate::bundle::Bundle::compile_source_ast_certified(
                language, &path, &module, &vm.modules, None, None, None,
                crate::frontend_preparation::snapshot(language.parse).argument_selection()).expect("emit prepared AST");
            assert_eq!(control.chunks.len(), direct.chunks.len(), "case {index}");
            for (left, right) in control.chunks.iter().zip(&direct.chunks) {
                assert_eq!(left.block_i32_results, right.block_i32_results);
                assert_eq!(left.i64_locals, right.i64_locals);
                assert_eq!(left.call_indirect_sigs, right.call_indirect_sigs);
                assert_eq!(left.call_indirect_canon, right.call_indirect_canon);
                assert_eq!(left.type_rec_shape, right.type_rec_shape);
                // Compare unordered metadata structurally above, so independent
                // hash seeds cannot make an otherwise exact result fail.
                let (mut left, mut right) = (left.clone(), right.clone());
                for chunk in [&mut left, &mut right] {
                    chunk.block_i32_results.clear();
                    chunk.i64_locals.clear();
                    chunk.call_indirect_sigs.clear();
                    chunk.call_indirect_canon.clear();
                    chunk.type_rec_shape.clear();
                }
                assert_eq!(format!("{left:?}"), format!("{right:?}"), "case {index}, chunk {}", left.name);
            }
            assert_eq!(vybe_platform_wasm::write_wasm(&control.chunks),
                vybe_platform_wasm::write_wasm(&direct.chunks), "case {index}");
            assert_eq!(format!("{:?}", (&control.required_libraries, &control.host_imports,
                control.component_source, &control.component_binary, &control.app_shell)),
                format!("{:?}", (&direct.required_libraries, &direct.host_imports,
                direct.component_source, &direct.component_binary, &direct.app_shell)), "case {index}");
        }
    }

    #[test]
    fn trait_preparation_without_uses_preserves_complete_module() {
        use vybe_ast::*;
        let declaration=Statement::new(StmtKind::ClassDecl{name:"T".into(),parents:vec![],interfaces:vec![],members:vec![],modifiers:ClassModifiers{kind:ClassKind::Trait,..Default::default()},decorators:vec![]});
        let mut runtime=super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
        runtime.php_trait_declarations.insert("t".into(),declaration.clone());
        for body in [vec![declaration.clone()],vec![Statement::new(StmtKind::NamespaceDecl{name:"N".into(),body:vec![Statement::new(StmtKind::Block(vec![declaration.clone()]))]})]] {
            let mut module=Module{name:"trait-no-uses".into(),language:Lang::PHP,body,imports:vec![],directives:Default::default(),canon:Default::default()};
            let before=format!("{module:?}");runtime.include_available_php_traits(&mut module);assert_eq!(format!("{module:?}"),before);
        }
    }

    #[test]
    fn local_trait_membership_preserves_kind_case_and_forward_visibility() {
        use vybe_ast::*;
        fn declaration(name: &str, kind: ClassKind, source: Option<&str>) -> Statement {
            Statement::new(StmtKind::ClassDecl {
                name: name.into(), parents: vec![], interfaces: vec![],
                members: source.into_iter().map(|from| ClassMember::Augment(AugmentDecl {
                    from: from.into(), ..Default::default()
                })).collect(),
                modifiers: ClassModifiers { kind, ..Default::default() }, decorators: vec![],
            })
        }
        for (name, used, local_kind, expected_insert) in [
            ("N.T", "n.t", ClassKind::Trait, false),
            ("N.T", "n.t", ClassKind::Class, true),
            ("ÉT", "Ét", ClassKind::Trait, false),
            ("ÉT", "ét", ClassKind::Trait, true),
        ] {
            let mut runtime = super::PhpIncludeRuntime::new(vybe_runtime::capabilities::Capabilities::all());
            runtime.php_trait_declarations.insert(used.to_ascii_lowercase(),
                declaration("Imported.T", ClassKind::Trait, None));
            let mut module = Module {
                name: "local-trait-membership".into(), language: Lang::PHP,
                body: vec![declaration("C", ClassKind::Class, Some(used)),
                    Statement::new(StmtKind::NamespaceDecl {
                        name: "Namespace".into(), body: vec![Statement::new(StmtKind::Block(
                            vec![declaration(name, local_kind, None)]))],
                    })],
                imports: vec![], directives: Default::default(), canon: Default::default(),
            };
            let before = format!("{module:?}");
            runtime.include_available_php_traits(&mut module);
            if expected_insert {
                assert_eq!(module.body.len(), 3);
                assert!(matches!(&module.body[0].kind, StmtKind::ClassDecl { name, .. } if name == "Imported.T"));
            } else {
                assert_eq!(format!("{module:?}"), before);
            }
        }
    }

}

#[cfg(v6_service_gate)]
#[path="retirement_service_gate.rs"]
pub mod retirement_service_gate;

#[cfg(any(test, feature="service-verification"))]
#[path="dynamic/retired_verification.rs"]
mod retired_verification;

/// Execute native request-local reuse checks after registering the real PHP frontend.
#[cfg(feature="service-verification")]
pub fn verify_retired_compiled_reuse() {
    retired_verification::retired_compiled_service_hit_moves_code_releases_metadata_and_consumes_once();
    retired_verification::retired_compiled_service_revalidates_exact_source_and_negative_library_fact();
    retired_verification::retired_compiled_count_payload_bounds_and_cancellation_release();
    retired_verification::retired_compiled_errors_are_not_retained_or_published();
}

#[cfg(feature="service-verification")]
#[path="dynamic/deferred_class_verification.rs"]
mod deferred_class_verification;
#[cfg(feature="service-verification")]
pub fn verify_deferred_class_donation_service(vm: VM) { deferred_class_verification::verify(vm); }

#[cfg(feature="service-verification")]mod request_code_space_verification;
#[cfg(feature="service-verification")]pub fn verify_request_code_space_service(vm:VM){request_code_space_verification::verify(vm); }

#[cfg(feature="service-verification")]use fused_manifest_verification::{FusedManifestProbe,FUSED_MANIFEST_PROBE};
#[cfg(feature="service-verification")]mod fused_manifest_verification{
use super::*;
// Feature-only actual Compiler gate. No policy fields or branches remain in
// normal builds; normal registered owned production always borrows its manifest.
#[derive(Clone,Copy,Debug,Default)]
pub struct FusedManifestCounts {pub borrowed:usize,pub clones:usize,pub accepted:usize,pub completed:usize}
pub(super) struct FusedManifestProbe {force_clone:bool,borrowed:AtomicUsize,clones:AtomicUsize,pub(super) accepted:AtomicUsize,pub(super) completed:AtomicUsize}
impl FusedManifestProbe {
    fn new(force_clone:bool)->Arc<Self>{Arc::new(Self{force_clone,borrowed:AtomicUsize::new(0),clones:AtomicUsize::new(0),accepted:AtomicUsize::new(0),completed:AtomicUsize::new(0)})}
    pub(super) fn use_borrowed(&self)->bool{if self.force_clone{self.clones.fetch_add(1,Ordering::Relaxed);}else{self.borrowed.fetch_add(1,Ordering::Relaxed);}!self.force_clone}
    fn counts(&self)->FusedManifestCounts{FusedManifestCounts{borrowed:self.borrowed.load(Ordering::Relaxed),clones:self.clones.load(Ordering::Relaxed),accepted:self.accepted.load(Ordering::Relaxed),completed:self.completed.load(Ordering::Relaxed)}}
}
thread_local! {pub(super) static FUSED_MANIFEST_PROBE:RefCell<Option<Arc<FusedManifestProbe>>>=const{RefCell::new(None)};}
/// Scoped validation policy captured once into real worker contexts. Existing
/// contexts retain their captured selection. Nesting/unwinding restores state.
pub fn with_fused_manifest_probe<R>(force_clone:bool,run:impl FnOnce()->R)->(R,FusedManifestCounts){
    struct Restore(Option<Arc<FusedManifestProbe>>);impl Drop for Restore{fn drop(&mut self){FUSED_MANIFEST_PROBE.with(|slot|{slot.replace(self.0.take());});}}
    let probe=FusedManifestProbe::new(force_clone);let old=FUSED_MANIFEST_PROBE.with(|slot|slot.replace(Some(probe.clone())));let restore=Restore(old);let result=run();drop(restore);(result,probe.counts())
}
pub fn fused_manifest_probe_counts()->FusedManifestCounts{FUSED_MANIFEST_PROBE.with(|slot|slot.borrow().as_ref().map_or_else(FusedManifestCounts::default,|probe|probe.counts()))}
/// Owns the real prepared ticket and captured worker context. Preparation,
/// source parsing and context capture happen once outside the timed emit loop.
/// No frontend/schema API is added and no second whole AST is stored.
pub struct FusedManifestValidation {
    context:FusedLoweringContext,prepared:crate::preparation::PreparedFile,
    language:Language,budget:Arc<AtomicUsize>,probe:Arc<FusedManifestProbe>,
}
impl FusedManifestValidation {
    pub fn prepare(vm:&VM,path:&Path,force_clone:bool)->Result<Self,String>{
        let language=languages::find_by_name("php").ok_or("PHP missing")?;
        let registration=crate::frontend_preparation::snapshot(language.parse);
        let pool=crate::preparation::PreparationPool::new(1,1).map_err(|e|e.to_string())?;
        let prepared=pool.submit(path.canonicalize().map_err(|e|e.to_string())?,language).ok_or("preparation declined")?.wait().ok_or("preparation failed")?;
        if !prepared.statement_production.as_ref().is_some_and(|p|p.is_owned()){return Err("actual prepared producer is not owned".into());}
        let mut helpers=crate::primitives::helper_templates::Session::default();
        let mut context=FusedLoweringContext::capture(vm,language,|_|true,helpers.selected(registration.helper_selection()),registration)?;
        let probe=FusedManifestProbe::new(force_clone);context.manifest_probe=Some(probe.clone());
        Ok(Self{context,prepared,language,budget:Arc::new(AtomicUsize::new(0)),probe})
    }
    pub fn set_force_clone(&mut self,force_clone:bool){let probe=FusedManifestProbe::new(force_clone);self.context.manifest_probe=Some(probe.clone());self.probe=probe;}
    pub fn emit(&mut self)->Result<DynamicCompilation,String>{
        let lowered=self.context.lower_prepared(&mut self.prepared,self.language,&self.budget).ok_or("actual fused lowering declined")?;
        lowered.compilation.map(PendingCompilation::into_raw)
    }
    pub fn counts(&self)->FusedManifestCounts{self.probe.counts()}
    pub fn source_state(&self)->Result<(usize,u64),String>{
        use std::hash::Hasher;
        struct Digest(std::hash::DefaultHasher);impl std::fmt::Write for Digest{fn write_str(&mut self,s:&str)->std::fmt::Result{self.0.write(s.as_bytes());Ok(())}}
        let module=self.prepared.module.as_ref().map_err(Clone::clone)?;
        let mut digest=Digest(std::hash::DefaultHasher::new());std::fmt::write(&mut digest,format_args!("{module:?}")).unwrap();
        Ok((module.body.as_ptr()as usize,digest.0.finish()))
    }
}

}
#[cfg(feature="service-verification")]pub use fused_manifest_verification::{FusedManifestValidation,FusedManifestCounts,with_fused_manifest_probe,fused_manifest_probe_counts};

#[cfg(feature="service-verification")]pub use request_code_space::{ResidentEntryCloneCounts,with_resident_entry_clone_control};
