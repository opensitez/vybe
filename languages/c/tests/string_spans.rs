use std::path::Path;
use vybe_runtime::VM;

#[test]
fn linear_single_byte_spans_match_native_c() {
    let source = r#"#include <stdlib.h>
#include <string.h>
int main(void) {
    char *p=malloc(65);
    for(int length=0;length<=64;length++) {
        memset(p,'a',length);p[length]=0;
        if(strspn(p,"a")!=length) abort();
        if(strcspn(p,"z")!=length) abort();
        if(strpbrk(p,"z")!=0) abort();
        for(int hit=0;hit<length;hit++) {
            p[hit]='z';
            if(strspn(p,"a")!=hit) abort();
            if(strcspn(p,"z")!=hit) abort();
            if(strpbrk(p,"z")!=p+hit) abort();
            p[hit]='a';
        }
    }
    memset(p,'a',32);p[32]=0;p[17]='b';
    if(strspn(p,"a\0b")!=17) abort();
    if(strcspn(p,"b\0a")!=17) abort();
    if(strpbrk(p,"b\0a")!=p+17) abort();
    for(int length=0;length<=64;length++) {
        memset(p,' ',length);p[length]=0;
        if(strspn(p," \t\r\n")!=length) abort();
        if(strcspn(p,",;:")!=length) abort();
        if(strpbrk(p,",;:")!=0) abort();
        for(int hit=0;hit<length;hit++) {
            p[hit]=';';
            if(strspn(p," \t\r\n")!=hit) abort();
            if(strcspn(p,",;:")!=hit) abort();
            if(strpbrk(p,",;:")!=p+hit) abort();
            p[hit]=' ';
        }
    }
    memset(p,'a',32);p[32]=0;p[17]='c';
    if(strspn(p,"ab\0c")!=17) abort();
    if(strcspn(p,"cd\0a")!=17) abort();
    if(strpbrk(p,"cd\0a")!=p+17) abort();
    for(int byte=0;byte<256;byte++) {
        memset(p,byte,32);p[32]=0;
        int member=byte>='0' && byte<='7';
        if(strspn(p,"01234567")!=(member?32:0)) abort();
        if(strcspn(p,"01234567")!=((byte==0||member)?0:32)) abort();
        if(strpbrk(p,"01234567")!=(member?p:0)) abort();
    }
    free(p);return 0;
}"#;
    let folder = std::env::temp_dir().join(format!("vybe-span-reference-{}", std::process::id()));
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
        .compile_source_by_name(source, "c", Path::new("spans.c"))
        .unwrap();
    for name in ["__c_strspn_h", "__c_strcspn_h", "__c_strpbrk_h"] {
        assert!(!compiled.chunks.iter().any(|c| c.name == name));
    }
    let native = vybe_platform_wasm::write_wasm(&compiled.chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&native)
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
