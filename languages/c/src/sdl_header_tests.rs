#[test]
fn messagebox_constants_are_module_definitions_after_standard_wasm_reload() {
    let profile = std::sync::Arc::new(
        vybe_compiler::profile::parse_profile(crate::profile_source()).unwrap());
    for header in ["SDL.h", "SDL2/SDL.h", "SDL_messagebox.h", "SDL2/SDL_messagebox.h"] {
        let source = format!("#include <{header}>\nint main(void){{return \
            SDL_MESSAGEBOX_ERROR!=16 || SDL_MESSAGEBOX_INFORMATION!=64 \
            || SDL_MESSAGEBOX_COLOR_MAX!=5;}}");
        for parse in [crate::parse_with_peg, crate::parse_optimized] {
            let module = parse(&source).unwrap();
            let chunks = vybe_compiler::primitives::Compiler::with_shared_profile(profile.clone())
                .compile_with_imports(&module).unwrap().chunks;
            let bytes = vybe_platform_wasm::write_wasm_with_options(&chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false, include_name_section: true,
                });
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes).unwrap();
            let decoded = vybe_platform_wasm::read_wasm(&bytes).unwrap();
            assert!(!decoded.iter().flat_map(|chunk| &chunk.global_imports)
                .any(|import| import.name.starts_with("SDL_MESSAGEBOX_")), "{header}");
            let mut vm = vybe_runtime::VM::new();
            vybe_runtime::init_all_registered(&mut vm,
                &vybe_runtime::capabilities::Capabilities::all());
            vm.run(decoded).unwrap();
            assert_eq!(vm.pending_exit_code, 0, "{header}");
        }
    }
}
