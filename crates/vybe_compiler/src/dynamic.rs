use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

use crate::bundle::{Bundle, CompiledBundle, EntryPoint, SourceFile};
use crate::languages::{self, Language};
use crate::primitives::HostImportMetadata;
use vybe_runtime::capabilities::{Capabilities, Capability};
use vybe_runtime::chunk::Chunk;
use vybe_runtime::chunk::Import;
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{HostContext, ImportTarget, VM, Value};

thread_local! {
    static ACTIVE_PHP_RUNTIME: RefCell<Option<*mut PhpIncludeRuntime>> = const { RefCell::new(None) };
    static ACTIVE_JS_RUNTIME: RefCell<Option<*mut JsDynamicRuntime>> = const { RefCell::new(None) };
    static CONSTRUCTED_JS_FUNCTIONS: RefCell<HashMap<u64, *mut ConstructedJsFunction>> = RefCell::new(HashMap::new());
}

static NEXT_JS_DYNAMIC_ID: AtomicU64 = AtomicU64::new(1);

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

#[derive(Debug)]
pub struct DynamicCompilation {
    pub chunks: Vec<Chunk>,
    pub host_imports: HostImportMetadata,
    pub entry_path: Option<PathBuf>,
    /// What the entry module DECLARED about presenting a UI — see
    /// [`vybe_ast::Directives::app_shell`]. `None` states nothing, and the
    /// document answers instead.
    pub app_shell: Option<vybe_ast::AppShell>,
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
    fn get(&self, path: &Path, fingerprint: u64) -> Option<Result<DynamicCompilation, String>>;
    fn store(
        &self,
        path: &Path,
        fingerprint: u64,
        deps: Option<Vec<PathBuf>>,
        result: Result<&DynamicCompilation, &str>,
    );
}

/// Everything a runtime include's compile depends on BESIDES its own source.
///
/// `compile_full_with_modules_and_php_entry_override` reads two things from the
/// live VM: the module map (`modules.get(m).exports.get(n)`, so record CONTENT,
/// not just which modules exist) and the entry path, which PHP's magic
/// constants make observable. A cached compilation may only be reused where
/// both are what they were.
///
/// That map is not frozen during a run: `host_imports::install` runs on every
/// dynamic include and registers host functions, and registration is the one
/// path that writes `vm.modules` (`insert_host_module_export`). So the
/// fingerprint is checked per include rather than assumed stable for the run.
///
/// Order-independent (wrapping sum of per-entry hashes) so no sort is needed —
/// this runs on every include and has to stay far below the ~15-25ms compile it
/// is there to avoid.
pub fn module_fingerprint(vm: &VM, entry_path: &Path) -> u64 {
    use std::hash::{Hash, Hasher};

    fn hash_of(value: impl Hash) -> u64 {
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        value.hash(&mut hasher);
        hasher.finish()
    }

    let mut total: u64 = hash_of(entry_path);
    total = total.wrapping_mul(31).wrapping_add(vm.modules.len() as u64);
    for (name, record) in &vm.modules {
        let mut per_module = hash_of(name).wrapping_add(record.exports.len() as u64);
        for (export, entry) in &record.exports {
            // The export's IDENTITY, not just its name: re-registering the same
            // `(module, name)` with a different index leaves both counts equal
            // and would otherwise read as "nothing changed".
            per_module = per_module
                .wrapping_add(hash_of(export))
                .wrapping_add(hash_of(std::mem::discriminant(entry)))
                .wrapping_add(match entry {
                    vybe_runtime::ExportEntry::Function { idx } => *idx as u64,
                    vybe_runtime::ExportEntry::ResourceType { type_id } => *type_id as u64,
                    // A `Value` export re-registered under the same name with a
                    // different value is invisible here — only its variant is.
                    // Those are boot infrastructure (`ecma:math.PI`), written
                    // before any include runs, not something a running script
                    // rewrites.
                    _ => 0,
                });
        }
        total = total.wrapping_add(per_module);
    }
    total
}

pub fn run_with_js_dynamic_runtime(
    vm: &mut VM,
    caps: Capabilities,
    chunks: Vec<Chunk>,
) -> Result<Value, String> {
    ensure_js_runtime_registered(vm);
    let active_imports = chunks
        .first()
        .map(|c| c.imports.clone())
        .unwrap_or_default();
    let active_resolved_imports = resolve_imports(vm, &active_imports).unwrap_or_default();
    let mut js_runtime = JsDynamicRuntime::new(caps);
    let _guard = js_runtime.activate(vm, active_imports, active_resolved_imports);
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
    ensure_js_runtime_registered(vm);
}

struct PhpIncludeRuntime {
    caps: Capabilities,
    vm: *mut VM,
    current_paths: Vec<PathBuf>,
    /// Per-RUN, and it stays that way. Only the COMPILATION is cacheable: share
    /// this across requests and request two sees request one's "already
    /// included" and skips the include entirely.
    included_once: HashSet<PathBuf>,
    active_imports: Vec<Import>,
    active_resolved_imports: Vec<ImportTarget>,
    include_cache: Option<std::sync::Arc<dyn IncludeCompileCache>>,
}

struct ActivePhpRuntimeGuard {
    previous: Option<*mut PhpIncludeRuntime>,
}

struct JsDynamicRuntime {
    caps: Capabilities,
    vm: *mut VM,
    active_imports: Vec<Import>,
    active_resolved_imports: Vec<ImportTarget>,
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

    /// Give runtime includes somewhere to reuse a compilation from. Without
    /// one they compile from source every time, which is the behaviour this
    /// service has always had.
    pub fn set_include_cache(&mut self, cache: std::sync::Arc<dyn IncludeCompileCache>) {
        self.php_runtime.include_cache = Some(cache);
    }

    pub fn compile_bundle(&mut self, bundle: &Bundle) -> Result<DynamicCompilation, String> {
        ensure_php_runtime_registered(self.vm);
        ensure_js_runtime_registered(self.vm);
        let compiled = bundle.compile_full_with_modules(&self.vm.modules)?;
        Ok(DynamicCompilation {
            chunks: compiled.chunks,
            host_imports: compiled.host_imports,
            entry_path: bundle.sources.first().map(|source| source.path.clone()),
            app_shell: compiled.app_shell,
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
        ensure_php_runtime_registered(self.vm);
        ensure_js_runtime_registered(self.vm);
        let base_chunk_index = self.vm.chunks.len();
        crate::host_imports::install(self.vm, &compiled.host_imports);
        install_chunk_globals(self.vm, &compiled.chunks, base_chunk_index);
        let active_imports = compiled
            .chunks
            .first()
            .map(|chunk| chunk.imports.clone())
            .unwrap_or_default();
        let active_resolved_imports = resolve_imports(self.vm, &active_imports)?;
        let _php_runtime = self.php_runtime.activate(
            self.vm,
            compiled.entry_path.as_deref(),
            active_imports.clone(),
            active_resolved_imports.clone(),
        );
        let _js_runtime =
            self.js_runtime
                .activate(self.vm, active_imports, active_resolved_imports.clone());
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
    //    Register WITH gui so the gui module's (module,name) keys exist here to be
    //    overlaid below — otherwise `vybe.gui.*` in an eval expression would fail
    //    to link. The fresh GuiState this creates is immediately shadowed by the
    //    live closures in `overlay_host_fns_from`.
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

/// A value that can be lifted into the eval mini-VM. Function/host-function
/// values can't cross VMs (their chunk_index refs are VM-local), so they are
/// excluded.
fn eval_value_is_copyable(_v: &Value) -> bool {
    // Copy EVERY global into the eval mini-VM, including function/constructor
    // values. They are Arc-shared, so property READS (`typeof C.prototype.x`,
    // walking a class's prototype chain) see live state — essential for
    // debugger inspection of classes. Calling a copied user `Function` in the
    // mini-VM would dispatch on the live VM's chunk_index (wrong table) and
    // error, but that degrades gracefully; `HostFunction(idx)` stays call-safe
    // because register_all is deterministic (same index in the mini-VM).
    true
}

pub fn install_chunk_globals(vm: &mut VM, chunks: &[Chunk], base_chunk_index: usize) {
    for (idx, chunk) in chunks.iter().enumerate() {
        if !should_publish_chunk_name(&chunk.name) {
            continue;
        }

        let func = Function {
            name: Some(chunk.name.clone()),
            arity: chunk.arity,
            chunk_index: base_chunk_index + idx,
            upvalues: vec![],
        };
        let mut obj = Object::new();
        obj.kind = ObjectKind::Function(func);
        let val = Value::Object(vybe_runtime::heap::alloc(obj));
        vm.set_global_owned(chunk.name.to_lowercase(), val);
    }
}

fn should_publish_chunk_name(name: &str) -> bool {
    !name.is_empty()
        && name != "<script>"
        && name != "<bootstrap>"
        && !name.starts_with("__stdlib_")
}

fn preprocess_dynamic_php_source(source: &str, source_path: &Path) -> String {
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
        chunks: compiled.chunks,
        host_imports: compiled.host_imports,
        entry_path: None,
        app_shell: compiled.app_shell,
    }
}

impl PhpIncludeRuntime {
    fn new(caps: Capabilities) -> Self {
        Self {
            caps,
            vm: std::ptr::null_mut(),
            current_paths: Vec::new(),
            included_once: HashSet::new(),
            include_cache: None,
            active_imports: Vec::new(),
            active_resolved_imports: Vec::new(),
        }
    }

    fn activate(
        &mut self,
        vm: &mut VM,
        entry_path: Option<&Path>,
        active_imports: Vec<Import>,
        active_resolved_imports: Vec<ImportTarget>,
    ) -> ActivePhpRuntimeGuard {
        self.vm = vm as *mut VM;
        self.current_paths.clear();
        if let Some(path) = entry_path {
            self.current_paths.push(path.to_path_buf());
        }
        self.active_imports = active_imports;
        self.active_resolved_imports = active_resolved_imports;
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

        if matches!(kind.as_str(), "include_once" | "require_once")
            && self.included_once.contains(&canonical_path)
        {
            return Ok(Value::I32(1));
        }

        let source = match fs::read_to_string(&resolved_path) {
            Ok(source) => source,
            Err(e) => {
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

        let language = languages::find_by_name("php")
            .ok_or_else(|| "php language profile missing".to_string())?;

        // Reuse an earlier compilation of this exact file only where everything
        // the compile READ is still what it was — the file itself and whatever
        // it statically pulls in (the dependency set), plus the VM's module map
        // and the entry path (the fingerprint).
        let cache = self.include_cache.clone();
        let fingerprint = cache
            .as_ref()
            .map(|_| module_fingerprint(vm, &entry))
            .unwrap_or(0);
        if let Some(cache) = cache.as_ref() {
            if let Some(hit) = cache.get(&canonical_path, fingerprint) {
                // A cached compile FAILURE is still that file's answer; letting
                // it fall through would recompile a broken include on every
                // request, which is exactly what the entry cache refuses to do.
                let compiled = hit?;
                return self.run_included_compilation(vm, compiled, kind, canonical_path);
            }
        }

        let bundle = bundle_from_source(source, language, resolved_path.clone());
        let compiled = match cache.as_ref() {
            Some(cache) => {
                // The dependency set has to be collected DURING the compile —
                // `prepared_module` opens files this bundle never names. The
                // recorder returns `None` when an outer scope owns the reads,
                // and `store` then caches nothing rather than risk staleness.
                let (result, deps) = crate::bundle::record_source_reads(|| {
                    self.compile_dynamic_php(vm, &bundle, &entry)
                });
                let deps = deps.map(|mut deps| {
                    // The include's own source is read by THIS function, before
                    // the recorded scope opens, so it is never in `deps` — and
                    // it is the one file whose edit must invalidate.
                    deps.push(resolved_path.clone());
                    deps
                });
                cache.store(
                    &canonical_path,
                    fingerprint,
                    deps,
                    result.as_ref().map_err(|err| err.as_str()),
                );
                result?
            }
            None => self.compile_dynamic_php(vm, &bundle, &entry)?,
        };
        self.run_included_compilation(vm, compiled, kind, canonical_path)
    }

    /// Install a compiled include into the live VM and run it — the half that
    /// is identical whether the compilation was just produced or came from the
    /// cache.
    fn run_included_compilation(
        &mut self,
        vm: &mut VM,
        compiled: DynamicCompilation,
        kind: String,
        canonical_path: PathBuf,
    ) -> Result<Value, String> {
        let base_chunk_index = vm.chunks.len();
        crate::host_imports::install(vm, &compiled.host_imports);
        install_chunk_globals(vm, &compiled.chunks, base_chunk_index);

        let child_active_imports = compiled
            .chunks
            .first()
            .map(|chunk| chunk.imports.clone())
            .unwrap_or_default();
        let child_active_resolved_imports = resolve_imports(vm, &child_active_imports)?;
        let saved_active_imports =
            std::mem::replace(&mut self.active_imports, child_active_imports);
        let saved_active_resolved_imports = std::mem::replace(
            &mut self.active_resolved_imports,
            child_active_resolved_imports.clone(),
        );

        self.current_paths.push(canonical_path.clone());
        let result = vm
            .run_linked_nested(compiled.chunks, child_active_resolved_imports)
            .map_err(|err| err.to_string());
        self.current_paths.pop();
        self.active_imports = saved_active_imports;
        self.active_resolved_imports = saved_active_resolved_imports;

        match result {
            Ok(Value::Null) => {
                if matches!(kind.as_str(), "include_once" | "require_once") {
                    self.included_once.insert(canonical_path);
                }
                Ok(Value::I32(1))
            }
            Ok(value) => {
                if matches!(kind.as_str(), "include_once" | "require_once") {
                    self.included_once.insert(canonical_path);
                }
                Ok(value)
            }
            // The included file ran and failed. That is the script's error, not
            // a failed file open — reporting it as `false` made a fatal inside
            // an include indistinguishable from a missing file.
            Err(err) => Err(err),
        }
    }

    fn compile_dynamic_php(
        &self,
        vm: &VM,
        bundle: &Bundle,
        entry_path: &Path,
    ) -> Result<DynamicCompilation, String> {
        let compiled = bundle
            .compile_full_with_modules_and_php_entry_override(&vm.modules, Some(entry_path))?;
        Ok(DynamicCompilation {
            chunks: compiled.chunks,
            host_imports: compiled.host_imports,
            entry_path: bundle.sources.first().map(|source| source.path.clone()),
            app_shell: compiled.app_shell,
        })
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
    ) -> Value {
        let vm = unsafe { &mut *self.vm };
        let compiled = match bundle.compile_full_with_modules(&vm.modules) {
            Ok(compiled) => compiled,
            Err(e) => return throw_eval_error(ctx, "SyntaxError", &e),
        };

        let base_chunk_index = vm.chunks.len();
        crate::host_imports::install(vm, &compiled.host_imports);
        install_chunk_globals(vm, &compiled.chunks, base_chunk_index);

        let child_active_imports = compiled
            .chunks
            .first()
            .map(|chunk| chunk.imports.clone())
            .unwrap_or_default();
        let child_active_resolved_imports = match resolve_imports(vm, &child_active_imports) {
            Ok(resolved) => resolved,
            Err(e) => return throw_eval_error(ctx, "EvalError", &e.to_string()),
        };
        let saved_active_imports =
            std::mem::replace(&mut self.active_imports, child_active_imports);
        let saved_active_resolved_imports = std::mem::replace(
            &mut self.active_resolved_imports,
            child_active_resolved_imports.clone(),
        );

        let result = vm.run_linked_nested(compiled.chunks, child_active_resolved_imports);

        self.active_imports = saved_active_imports;
        self.active_resolved_imports = saved_active_resolved_imports;

        let run_value = match result {
            Ok(value) => value,
            Err(e) => return throw_eval_error(ctx, "SyntaxError", &e.to_string()),
        };

        // An uncaught throw inside eval'd code propagates to the caller.
        if let Some(exc) = vm.last_exception.take() {
            ctx.throw_value(exc);
            return Value::Undefined;
        }

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
            active_imports: Vec::new(),
            active_resolved_imports: Vec::new(),
        }
    }

    fn activate(
        &mut self,
        vm: &mut VM,
        active_imports: Vec<Import>,
        active_resolved_imports: Vec<ImportTarget>,
    ) -> ActiveJsRuntimeGuard {
        self.vm = vm as *mut VM;
        self.active_imports = active_imports;
        self.active_resolved_imports = active_resolved_imports;
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

        // Per-language eval quirks:
        //  - PHP: the string is evaluated in `<?php` context (bare text is
        //    literal output, not code); its top-level `return` is the result.
        //  - Python: `eval` (attrs.completion_value == true) yields the value of
        //    the expression; `exec` yields None. Capture the expression value by
        //    binding it to a temp we read back after the run.
        let mut completion_capture: Option<&'static str> = None;
        let eval_source = match language_name.as_str() {
            "php" => {
                if source.trim_start().starts_with("<?") {
                    source.clone()
                } else {
                    format!("<?php {source}")
                }
            }
            "python" => {
                let wants_value = args
                    .get(2)
                    .map(|a| object_bool_prop(a, "completion_value"))
                    .unwrap_or(false);
                if wants_value {
                    completion_capture = Some("__vybe_eval_result__");
                    format!("__vybe_eval_result__ = ({})", source.trim())
                } else {
                    source.clone()
                }
            }
            _ => source.clone(),
        };

        let bundle = bundle_from_source(eval_source, language, PathBuf::from("<eval>"));

        // No explicit namespace dict → compile into the LIVE VM, exactly as
        // `handle_dynamic_include` does.
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
        // isolation by definition rather than caller-global mutation.
        // Python's expression `eval()` also stays on the mini-VM: it wants a
        // VALUE, and its result may be an object whose repr helper resolves
        // against the chunk table it was built in. Definitions arrive through
        // `exec()`, which sets no completion capture — so the live-VM route
        // covers exactly the cases that need to define something.
        let has_namespace_dict = args
            .get(2)
            .and_then(|a| object_get_prop(a, "namespace"))
            .is_some_and(|v| matches!(v, Value::Object(_)));
        if !has_namespace_dict && completion_capture.is_none() {
            return self.eval_in_live_vm(ctx, &bundle, completion_capture);
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

        let run_result = {
            let mut service =
                RuntimeCompilerService::with_capabilities(&mut eval_vm, self.caps.clone());
            match service.compile_and_run_bundle(&bundle) {
                Ok(value) => value,
                Err(e) => return throw_eval_error(ctx, "SyntaxError", &e),
            }
        };

        // Propagate an uncaught throw from eval'd code to the caller.
        if let Some(exc) = eval_vm.last_exception.take() {
            ctx.throw_value(exc);
            return Value::Undefined;
        }

        // Completion value: Python `eval` reads the captured temp; PHP (and any
        // language whose eval'd code uses an explicit top-level `return`) uses
        // the script's return value; `exec`/others return that value too.
        let result = match completion_capture {
            Some(name) => eval_vm.global(name).cloned().unwrap_or(Value::Undefined),
            None => run_result,
        };

        // Definitions/assignments escape: into the explicit namespace dict if
        // one was given, otherwise the caller's globals. (Skip the completion
        // temp; skip function values.) For the namespace, only write updates to
        // its own keys plus brand-new bindings — never the host builtins that
        // seeding left in `base_keys`.
        match &namespace {
            Some(ns) => {
                let mut guard = ns.lock().unwrap();
                for (k, v) in eval_vm.globals_by_name() {
                    if Some(k.as_str()) == completion_capture {
                        continue;
                    }
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
                    if Some(k.as_str()) == completion_capture {
                        continue;
                    }
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

                    let saved_this = state.vm.global("__js_this").cloned();
                    state.vm.set_global("__js_this", ctx.current_js_this());
                    ensure_js_runtime_registered(&mut state.vm);
                    let mut nested_runtime = JsDynamicRuntime::new(state.caps.clone());
                    let _guard = nested_runtime.activate(&mut state.vm, Vec::new(), Vec::new());
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

                    if let Some(saved_this) = saved_this {
                        state.vm.set_global("__js_this", saved_this);
                    } else {
                        state.vm.remove_global("__js_this");
                    }

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
    let key = ("vybe:php".to_string(), "dynamic_include".to_string());
    if vm.host_registry.contains_key(&key) {
        return;
    }

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
                    Err(message) => throw_dynamic_compile_error(ctx, message),
                }
            })
        }),
    );
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
    use std::path::PathBuf;
    use std::sync::{Arc, Mutex};
    use std::time::{SystemTime, UNIX_EPOCH};

    use super::{JsDynamicRuntime, RuntimeCompilerService, ensure_js_runtime_registered};
    use vybe_runtime::{VM, Value};

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
        let call_greet = vm
            .global("callgreet")
            .cloned()
            .expect("callGreet global");

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
            let _guard = runtime.activate(&mut vm, vec![], vec![]);
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
            let _guard = runtime.activate(&mut vm, vec![], vec![]);
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
            "<?php $path = 'partial.php'; $value = include $path; $result = get_value() + $value;",
        )
        .expect("write main php");
        std::fs::write(
            &partial_path,
            "<?php function get_value() { return 8; } return 34;",
        )
        .expect("write partial php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        service
            .compile_and_run_path(&main_path)
            .expect("run php with dynamic include");

        match vm.global("__php_var_result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected $result global, got {other:?}"),
        }
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
            "<?php $path = 'child.php'; $result = include $path;",
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
        service
            .compile_and_run_path(&main_path)
            .expect("run nested php dynamic include");

        match vm.global("__php_var_result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected nested $result global, got {other:?}"),
        }
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
            "<?php $result = require __DIR__ . '/compat.php'; $call_result = compat_polyfill();",
        )
        .expect("write main php");
        std::fs::write(
            &compat_path,
            "<?php\nif ( ! function_exists( 'compat_polyfill' ) ) :\n\t/**\n\t * Timing attack safe string comparison.\n\t *\n\t * @param string $known_string Expected string.\n\t * @param string $user_string  Actual, user supplied, string.\n\t * @return bool Whether strings are equal.\n\t */\n\tfunction compat_polyfill( $known_string = 'a', $user_string = 'a' ) {\n\t\t$known_string_length = strlen( $known_string );\n\n\t\tif ( strlen( $user_string ) !== $known_string_length ) {\n\t\t\treturn false;\n\t\t}\n\n\t\t$result = 0;\n\n\t\tfor ( $i = 0; $i < $known_string_length; $i++ ) {\n\t\t\t$result |= ord( $known_string[ $i ] ) ^ ord( $user_string[ $i ] );\n\t\t}\n\n\t\treturn 0 === $result ? 42 : 0;\n\t}\nendif;\n\nreturn compat_polyfill();\n",
        )
        .expect("write compat php");

        let mut vm = configured_vm();
        let mut service = RuntimeCompilerService::new(&mut vm);
        service
            .compile_and_run_path(&main_path)
            .expect("run php with alternative syntax include");

        match vm.global("__php_var_result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected include $result global, got {other:?}"),
        }

        match vm.global("__php_var_call_result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected include $call_result global, got {other:?}"),
        }
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
            "<?php $path = 'classes/child.php'; $result = include $path;",
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
        service
            .compile_and_run_path(&main_path)
            .expect("run php with entry-relative nested include");

        match vm.global("__php_var_result") {
            Some(Value::I32(value)) => assert_eq!(*value, 42),
            Some(Value::I64(value)) => assert_eq!(*value, 42),
            Some(Value::F64(value)) => assert_eq!(*value, 42.0),
            other => panic!("expected nested entry-relative $result global, got {other:?}"),
        }
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

        let mut vm = configured_vm();
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
}
