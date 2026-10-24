//! Export a compiled WAST caller for independent emitted-byte validation.
fn main() {
    let source = std::env::args().nth(1).expect("WAST source path");
    let output = std::env::args().nth(2).expect("output Wasm path");
    let module = vybe_language_wast::parse(&std::fs::read_to_string(source).unwrap()).unwrap();
    let profile = vybe_compiler::profile::parse_profile(vybe_language_wast::profile_source()).unwrap();
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module).unwrap();
    let bytes = vybe_platform_wasm::write_wasm(&chunks);
    std::fs::write(output, &bytes).unwrap();
    println!("{} bytes: {:?}", bytes.len(), vybe_platform_wasm::validate_binary(&bytes));
}
