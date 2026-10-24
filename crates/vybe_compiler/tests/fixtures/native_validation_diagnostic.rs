use vybe_platform_wasm::inspection::{CoreModule, ImportType};
use vybe_runtime::Chunk;

pub fn report(
    chunks: &[Chunk],
    bytes: &[u8],
    error: &vybe_platform_wasm::WasmError,
    before: usize,
    after: usize,
) {
    let named = vybe_platform_wasm::write_wasm_with_options(
        chunks,
        vybe_platform_wasm::WasmWriteOptions {
            include_name_section: true,
            ..Default::default()
        },
    );
    let mut names = std::collections::BTreeMap::new();
    CoreModule::new(&named)
        .unwrap()
        .function_names(|index, name| {
            names.insert(index, name.to_owned());
            Ok(())
        })
        .unwrap();
    let module = CoreModule::new(bytes).unwrap();
    let mut function = 0u32;
    module
        .imports(|_, _, ty| {
            if matches!(ty, ImportType::Func { .. }) {
                function += 1;
            }
            Ok(())
        })
        .unwrap();
    let Some(failure) = error.offset() else {
        eprintln!("native validation failure without instruction offset: {error}");
        return;
    };
    module
        .function_bodies(|body| {
            if body.range().contains(&failure) {
                eprintln!(
                    "native failure function={function} name={:?} range={:?} error={error}",
                    names.get(&function),
                    body.range()
                );
                for instruction in body.instructions()? {
                    let instruction = instruction?;
                    if instruction.offset.saturating_add(before) >= failure
                        && instruction.offset <= failure.saturating_add(after)
                    {
                        eprintln!("  wasm@{}: {instruction:?}", instruction.offset);
                        if matches!(instruction.opcode, 0x10 | 0x12) {
                            if let Some(index) = instruction.function_index() {
                                eprintln!("    target={:?}", names.get(&index));
                            }
                        }
                    }
                }
            }
            function += 1;
            Ok(())
        })
        .unwrap();
}
