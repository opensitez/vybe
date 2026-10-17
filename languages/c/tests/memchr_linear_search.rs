use std::path::Path;
use vybe_runtime::VM;

#[test]
fn unsigned_return_conversion_matches_generic_search() {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let source = "#include <stdlib.h>\n#include <string.h>\nunsigned int next_count(void){return 1.5;} int main(void){unsigned char *p=malloc(32); memset(p,145,32); p[1]=255; int old=__c_mem_index_of_h(p,255,next_count(),0); void *found=memchr(p,255,next_count()); if(old<0){if(found!=0) abort();}else{if(found!=p+old) abort();} free(p); return 0;}";
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", Path::new("unsigned_return_conversion.c"))
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn signed_count_initialization_matches_generic_search() {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let source = "#include <stdlib.h>\n#include <string.h>\nint main(void){unsigned char *p=malloc(32); memset(p,145,32); p[1]=255; int n=1.5; int old=__c_mem_index_of_h(p,255,n,0); void *found=memchr(p,255,n); if(old<0){if(found!=0) abort();}else{if(found!=p+old) abort();} free(p); return 0;}";
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", Path::new("signed_count_initialization.c"))
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn signed_negative_counts_do_not_read_memory_and_evaluate_once() {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let source = "#include <stdlib.h>\n#include <string.h>\nint count; int calls; int next_count(void){calls++; return count;} int main(void){unsigned char *p=malloc(32); memset(p,145,32); p[17]=255; int values[5]={-2147483648,-1,0,17,32}; for(int i=0;i<5;i++){count=values[i]; calls=0; void *found=memchr(p,255,next_count()); if(calls!=1) abort(); if(count<32){if(found!=0) abort();}else{if(found!=p+17) abort();}} int n=-1; if(memchr(p,255,n++)!=0||n!=0) abort(); free(p); return 0;}";
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", Path::new("signed_count_search.c"))
        .unwrap();
    assert!(
        compiled
            .chunks
            .iter()
            .any(|c| c.name == "__c_mem_index_of_h")
    );
    let wasm = vybe_platform_wasm::write_wasm(&compiled.chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn narrow_dynamic_needles_preserve_byte_wrapping() {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let source = "#include <stdlib.h>\n#include <string.h>\nint main(void){unsigned char *p=malloc(32); memset(p,145,32); p[17]=255; signed char a=-1; unsigned char b=255; short c=511; unsigned short d=511; if(memchr(p,a,32)!=p+17) abort(); if(memchr(p,b,32)!=p+17) abort(); if(memchr(p,c,32)!=p+17) abort(); if(memchr(p,d,32)!=p+17) abort(); free(p); return 0;}";
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", Path::new("narrow_needle_search.c"))
        .unwrap();
    assert_eq!(
        compiled
            .chunks
            .iter()
            .filter(|c| c.name == "<primitive-memory32-byte-index>")
            .count(),
        1
    );
    assert!(
        !compiled
            .chunks
            .iter()
            .any(|c| c.name == "__c_mem_index_of_h")
    );
    let wasm = vybe_platform_wasm::write_wasm(&compiled.chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn dynamic_integer_needles_wrap_and_evaluate_once() {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let source = "#include <stdlib.h>\n#include <string.h>\nint needle; int calls; int next_needle(void){calls++; return needle;} int main(void){unsigned char *p=malloc(32); int values[6]={0,255,511,-1,2147483647,-2147483648}; unsigned int n=32; for(int i=0;i<6;i++){memset(p,145,32); needle=values[i]; p[17]=(unsigned char)needle; calls=0; if(memchr(p,next_needle(),n)!=p+17) abort(); if(calls!=1) abort(); n=17; if(memchr(p,needle,n)!=0) abort(); n=32;} free(p); return 0;}";
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", Path::new("dynamic_needle_search.c"))
        .unwrap();
    assert_eq!(
        compiled
            .chunks
            .iter()
            .filter(|c| c.name == "<primitive-memory32-byte-index>")
            .count(),
        1
    );
    assert!(
        !compiled
            .chunks
            .iter()
            .any(|c| c.name == "__c_mem_index_of_h")
    );
    let wasm = vybe_platform_wasm::write_wasm(&compiled.chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn unsigned_dynamic_counts_keep_full_word_range_and_evaluate_once() {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let source = "#include <stdlib.h>\n#include <string.h>\nunsigned int count; int calls; unsigned int next_count(void){calls++; return count;} int main(void){unsigned char *p=malloc(32); memset(p,145,32); p[0]=255; unsigned int values[6]={0,1,17,32,0x80000000,0xffffffff}; for(int i=0;i<6;i++){count=values[i]; calls=0; void *found=memchr(p,255,next_count()); if(calls!=1) abort(); if(count==0){if(found!=0) abort();}else{if(found!=p) abort();}} unsigned int n=1; if(memchr(p,255,n++)!=p||n!=2) abort(); free(p); return 0;}";
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", Path::new("unsigned_count_search.c"))
        .unwrap();
    assert_eq!(
        compiled
            .chunks
            .iter()
            .filter(|c| c.name == "<primitive-memory32-byte-index>")
            .count(),
        1
    );
    assert!(
        !compiled
            .chunks
            .iter()
            .any(|c| c.name == "__c_mem_index_of_h")
    );
    let wasm = vybe_platform_wasm::write_wasm(&compiled.chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn proven_linear_search_preserves_bounds_zero_and_wrapped_needles() {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let source = "#include <stdlib.h>\n#include <string.h>\nint main(void){unsigned char *p=malloc(32); memset(p,145,32); p[17]=255; if(memchr(p,511,32)!=p+17) abort(); if(memchr(p,-1,32)!=p+17) abort(); if(memchr(p,255,17)!=0) abort(); p[0]=0; if(memchr(p,0,32)!=p) abort(); if(memchr(p,0,0)!=0) abort(); free(p); return 0;}";
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", Path::new("linear_search.c"))
        .unwrap();
    assert_eq!(
        compiled
            .chunks
            .iter()
            .filter(|c| c.name == "<primitive-memory32-byte-index>")
            .count(),
        1
    );
    assert!(
        !compiled
            .chunks
            .iter()
            .any(|c| c.name == "__c_mem_index_of_h")
    );
    let wasm = vybe_platform_wasm::write_wasm(&compiled.chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}

#[test]
fn signed_dynamic_count_retains_existing_search_recipe() {
    vybe_language_c::register();
    let mut vm = VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let source = "#include <stdlib.h>\n#include <string.h>\nint main(void){unsigned char *p=malloc(32); memset(p,145,32); p[17]=255; int n=32; if(memchr(p,255,n)!=p+17) abort(); n=17; if(memchr(p,255,n)!=0) abort(); free(p); return 0;}";
    let compiled = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .compile_source_by_name(source, "c", Path::new("dynamic_search.c"))
        .unwrap();
    assert!(
        compiled
            .chunks
            .iter()
            .any(|c| c.name == "__c_mem_index_of_h")
    );
    assert!(
        !compiled
            .chunks
            .iter()
            .any(|c| c.name == "<primitive-memory32-byte-index>")
    );
    let wasm = vybe_platform_wasm::write_wasm(&compiled.chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&wasm)
        .unwrap();
    vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm)
        .run_compiled(compiled)
        .unwrap();
    assert_eq!(vm.pending_exit_code, 0);
}
