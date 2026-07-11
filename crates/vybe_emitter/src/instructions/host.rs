use std::collections::HashMap;
use std::sync::OnceLock;

use vybe_bytecode::{Chunk, VM};

pub type EmitFn = fn(&FunctionRegistry, &mut Chunk, u32);

pub struct FunctionRegistry {
    functions: HashMap<(&'static str, &'static str), FunctionEntry>,
}

#[allow(dead_code)]
struct FunctionEntry {
    module: &'static str,
    name: &'static str,
}

impl FunctionRegistry {
    fn from_vm() -> Self {
        let mut vm = VM::new();
        vybe_host::register_all(&mut vm);
        let functions = vm
            .iter_host_function_exports()
            .map(|(module, name, _idx)| {
                let module: &'static str = Box::leak(module.into_boxed_str());
                let name: &'static str = Box::leak(name.into_boxed_str());
                ((module, name), FunctionEntry { module, name })
            })
            .collect();
        Self { functions }
    }

    pub fn emit(&self, c: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
        debug_assert!(
            self.functions
                .keys()
                .any(|(m, n)| *m == module && *n == name),
            "host function {module}.{name} was not registered by vybe_host::register_all"
        );
        let idx = c.add_import(module, name);
        c.emit_call(idx, argc, line);
    }

    pub fn has(&self, module: &str, name: &str) -> bool {
        self.functions
            .keys()
            .any(|(m, n)| *m == module && *n == name)
    }

    /// Every registered host export as `(module, name)` — the component-model
    /// interface surface. The namespace tree mounts from this
    /// (`namespaces::resolve_path` lazy mount), so resolution and emission
    /// share ONE source of truth for what the host exports.
    pub fn entries(&self) -> impl Iterator<Item = (&'static str, &'static str)> + '_ {
        self.functions.keys().copied()
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
