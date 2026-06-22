use std::collections::HashSet;
use std::sync::OnceLock;

use vybe_bytecode::{Chunk, VM};

static REGISTERED_HOST_FUNCTIONS: OnceLock<HashSet<(&'static str, &'static str)>> =
    OnceLock::new();

fn registered_host_functions() -> &'static HashSet<(&'static str, &'static str)> {
    REGISTERED_HOST_FUNCTIONS.get_or_init(|| {
        let mut vm = VM::new();
        vybe_host::register_all(&mut vm);
        vm.iter_host_function_exports()
            .map(|(module, name, _idx)| {
                let module: &'static str = Box::leak(module.into_boxed_str());
                let name: &'static str = Box::leak(name.into_boxed_str());
                (module, name)
            })
            .collect()
    })
}

pub fn emit(c: &mut Chunk, module: &'static str, name: &'static str, argc: u8, line: u32) {
    debug_assert!(
        registered_host_functions().contains(&(module, name)),
        "host function {module}.{name} was not registered by vybe_host::register_all"
    );
    let idx = c.add_import(module, name);
    c.emit_call(idx, argc, line);
}
