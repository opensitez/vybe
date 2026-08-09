use std::collections::{HashMap, HashSet};
use std::sync::OnceLock;

use vybe_runtime::{Chunk, VM};

pub type EmitFn = fn(&FunctionRegistry, &mut Chunk, u32);

/// Which host functions the platform plugins registered, as module → names.
///
/// Nested rather than keyed by `(module, name)` so a lookup is an actual hash
/// lookup. With a `(&'static str, &'static str)` key there is no way to probe
/// the map with a borrowed `(&str, &str)` — `Borrow` does not reach inside a
/// tuple — so both queries below degenerated into `.keys().any(...)`, a linear
/// scan of every host export (~1,900) on EVERY emitted host call and every
/// resolution probe. Nesting works because `&'static str: Borrow<str>`, so
/// `get(module)` and `contains(name)` take a plain `&str`.
pub struct FunctionRegistry {
    functions: HashMap<&'static str, HashSet<&'static str>>,
}

impl FunctionRegistry {
    fn from_vm() -> Self {
        // Emit-time validation only: enumerate the host functions the platform
        // plugins provide, so `emit` can assert a `module.name` really exists.
        // Runs the platform plugins through the one plugin loop — languages are
        // irrelevant here (they register descriptors, not host fns).
        let mut vm = VM::new();
        crate::primitives::platforms::register_platforms_all(&mut vm);
        let mut functions: HashMap<&'static str, HashSet<&'static str>> = HashMap::new();
        for (module, name, _idx) in vm.iter_host_function_exports() {
            let module: &'static str = Box::leak(module.into_boxed_str());
            let name: &'static str = Box::leak(name.into_boxed_str());
            functions.entry(module).or_default().insert(name);
        }
        Self { functions }
    }

    pub fn emit(&self, c: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
        debug_assert!(
            self.has(module, name),
            "host function {module}.{name} was not registered by the platform plugins"
        );
        let idx = c.add_import(module, name);
        c.emit_call(idx, argc, line);
    }

    pub fn has(&self, module: &str, name: &str) -> bool {
        self.functions
            .get(module)
            .is_some_and(|names| names.contains(name))
    }

    /// Every registered host export as `(module, name)` — the component-model
    /// interface surface. The namespace tree mounts from this
    /// (`namespaces::resolve_path` lazy mount), so resolution and emission
    /// share ONE source of truth for what the host exports.
    ///
    /// Unordered, as it was before: the one consumer mounts each pair into a
    /// tree, where order cannot matter.
    pub fn entries(&self) -> impl Iterator<Item = (&'static str, &'static str)> + '_ {
        self.functions
            .iter()
            .flat_map(|(module, names)| names.iter().map(move |name| (*module, *name)))
    }
}

pub struct EmitRegistry {
    recipes: HashMap<&'static str, EmitFn>,
}

impl EmitRegistry {
    pub fn new() -> Self {
        Self {
            recipes: HashMap::new(),
        }
    }

    pub fn register(&mut self, name: &'static str, f: EmitFn) {
        self.recipes.insert(name, f);
    }

    pub fn emit(&self, name: &str, fns: &FunctionRegistry, c: &mut Chunk, line: u32) -> bool {
        if let Some(f) = self.recipes.get(name) {
            f(fns, c, line);
            true
        } else {
            false
        }
    }
}

pub struct CapabilityContext {
    pub functions: FunctionRegistry,
    pub emits: EmitRegistry,
}

static CONTEXT: OnceLock<CapabilityContext> = OnceLock::new();

impl CapabilityContext {
    pub fn get() -> &'static Self {
        CONTEXT.get_or_init(|| {
            let functions = FunctionRegistry::from_vm();
            let mut emits = EmitRegistry::new();
            super::recipes::register_all(&functions, &mut emits);
            CapabilityContext { functions, emits }
        })
    }
}

pub fn emit(c: &mut Chunk, module: &'static str, name: &'static str, argc: u8, line: u32) {
    CapabilityContext::get()
        .functions
        .emit(c, module, name, argc, line);
}
