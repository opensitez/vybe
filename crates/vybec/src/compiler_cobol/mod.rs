pub mod compiler;

pub use compiler::Compiler;

use vybe_bytecode::VM;

/// Set up VM with all host functions needed by COBOL, then compile.
pub fn setup_and_compile(
    vm: &mut VM,
    program: &crate::parser_cobol::Program,
) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    vybe_host::register_all(vm);
    Compiler::new().compile(program)
}

/// Compile COBOL to a Component (cross-language module with exports).
pub fn compile_component(
    program: &crate::parser_cobol::Program,
    module_name: &str,
) -> Result<vybe_bytecode::component::Component, String> {
    let chunks = Compiler::new().compile(program)?;
    Ok(vybe_compiler_common::components::build_component(
        module_name,
        vybe_bytecode::component::Language::Cobol,
        chunks,
    ))
}
