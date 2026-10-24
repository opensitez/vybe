//! Compare complete native modules emitted from one real PHP compilation.
//! No guest execution or application timing claim is made by this size census.
use vybe_compiler::dynamic::RuntimeCompilerService;
use vybe_runtime::{VM, capabilities::Capabilities};
#[path = "../../../platforms/wasm/tests/fixtures/call_dispatch_baseline.rs"]
mod baseline;

fn bodies(bytes: &[u8]) -> Vec<usize> {
    let mut sizes = Vec::new();
    vybe_platform_wasm::inspection::CoreModule::new(bytes)
        .unwrap()
        .function_bodies(|body| {
            sizes.push(body.bytes().len());
            Ok(())
        })
        .unwrap();
    sizes
}

fn instruction_counts(bytes: &[u8]) -> Vec<usize> {
    let mut counts = Vec::new();
    vybe_platform_wasm::inspection::CoreModule::new(bytes)
        .unwrap()
        .function_bodies(|body| {
            let mut count = 0;
            for op in body.instructions()? {
                op?;
                count += 1;
            }
            counts.push(count);
            Ok(())
        })
        .unwrap();
    counts
}

fn main() {
    let path = std::path::PathBuf::from(std::env::args().nth(1).expect("PHP source path"));
    vybe_language_php::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &Capabilities::all());
    let compiled = RuntimeCompilerService::with_capabilities(&mut vm, Capabilities::all())
        .compile_path(&path)
        .unwrap();
    let chunks = &compiled.chunks;
    let shared = vybe_platform_wasm::write_wasm(chunks);
    let inline = baseline::inline_module(chunks, &shared);
    let mut validation = serde_json::Map::new();
    for (name, bytes) in [("inline", &inline), ("shared", &shared)] {
        let own = vybe_platform_wasm::validate_binary(bytes)
            .map(|_| "pass".to_owned())
            .unwrap_or_else(|e| e.to_string());
        let file = std::env::temp_dir().join(format!(
            "vybe-native-dispatch-{}-{name}.wasm",
            std::process::id()
        ));
        std::fs::write(&file, bytes).unwrap();
        let output = std::process::Command::new("node").args(["--experimental-wasm-custom-descriptors", "--experimental-wasm-wasmfx", "-e"])
            .arg("new WebAssembly.Module(require('node:fs').readFileSync(process.argv[1])); console.log('pass');")
            .arg(&file).output().unwrap();
        std::fs::remove_file(file).unwrap();
        let native = if output.status.success() {
            "pass".to_owned()
        } else {
            String::from_utf8_lossy(&output.stderr).into_owned()
        };
        validation.insert(name.into(), serde_json::json!({"reader":own,"v8":native}));
    }
    let before = bodies(&inline);
    let after = bodies(&shared);
    let before_ops = instruction_counts(&inline);
    let after_ops = instruction_counts(&shared);
    let mut functions: Vec<_> = chunks
        .iter()
        .enumerate()
        .map(|(index, c)| {
            serde_json::json!({
        "name":c.name,"source_index":index,"ir_bytes":c.code.len(),
        "before_body_bytes":before[index],"after_body_bytes":after[index],
        "before_native_instructions":before_ops[index],"after_native_instructions":after_ops[index]})
        })
        .collect();
    functions.sort_by_key(|f| std::cmp::Reverse(f["before_body_bytes"].as_u64().unwrap()));
    functions.truncate(20);
    println!(
        "{}",
        serde_json::json!({"source":path,"chunks":chunks.len(),
        "before_module_bytes":inline.len(),"after_module_bytes":shared.len(),
        "before_code_body_bytes":before.iter().sum::<usize>(),
        "after_code_body_bytes":after.iter().sum::<usize>(),
        "before_native_instructions":before_ops.iter().sum::<usize>(),
        "after_native_instructions":after_ops.iter().sum::<usize>(),
        "added_helpers":after.len()-before.len(),"functions":functions,
        "validation":validation,"execution_claim":false})
    );
}
