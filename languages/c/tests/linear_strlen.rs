use std::path::Path;
use vybe_runtime::VM;
#[test]
fn linear_strlen_keeps_every_terminator_position() {
    let source=r#"#include <stdlib.h>
#include <string.h>
int main(void){char *p=malloc(8193);for(int n=0;n<=96;n++){memset(p,145,n);p[n]=0;if(strlen(p)!=n)abort();}memset(p,145,8192);p[8192]=0;if(strlen(p)!=8192)abort();free(p);return 0;}
"#;
    let folder = std::env::temp_dir().join(format!("vybe-length-reference-{}", std::process::id()));
    std::fs::create_dir(&folder).unwrap();
    let path = folder.join("reference.c");
    let executable = folder.join("reference");
    std::fs::write(&path, source).unwrap();
    assert!(
        std::process::Command::new("cc")
            .arg("-O0")
            .arg(&path)
            .arg("-o")
            .arg(&executable)
            .status()
            .unwrap()
            .success()
    );
    assert!(
        std::process::Command::new(&executable)
            .status()
            .unwrap()
            .success()
    );
    std::fs::remove_file(executable).unwrap();
    std::fs::remove_file(path).unwrap();
    std::fs::remove_dir(folder).unwrap();
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", Path::new("length.c"))
        .unwrap();
    let native = vybe_platform_wasm::write_wasm(&compiled.chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&native)
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
