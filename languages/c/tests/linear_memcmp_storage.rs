use vybe_compiler::{Compiler, primitives};
use vybe_runtime::{Op, VM, profile::BuiltinEmit};
const SOURCE: &str = r#"
typedef unsigned long size_t;
void *malloc(size_t size);
int memcmp(const void *a, const void *b, size_t n);
int result;
int visits;
int main() {
  int *a=malloc(128); int *b=malloc(128);
  for(int i=0;i<32;i++){a[i]=i+17;b[i]=i+17;}
  result=memcmp((visits=visits*10+1,a),(visits=visits*10+2,b),128);
  if(result!=0)return 10;
  if(visits!=12)return 11;
  result=memcmp(a,a,128); if(result!=0)return 12;
  result=memcmp(a,b,0); if(result!=0)return 13;
  b[3]=b[3]+1;
  result=memcmp(a,b,128);
  if(result>=0)return 14;
  return 0;
}
"#;
fn compile(source: &str, fast: bool) -> Vec<vybe_runtime::Chunk> {
    vybe_language_c::register();
    let ast = vybe_language_c::parse(source).unwrap();
    let mut profile =
        vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap();
    if !fast {
        profile
            .builtins
            .get_mut("__c_ptr_byte_compare")
            .unwrap()
            .emit = BuiltinEmit::Common("c.memcmp".into());
    }
    Compiler::with_profile(profile).compile(&ast).unwrap()
}
fn vector_count(chunks: &[vybe_runtime::Chunk]) -> usize {
    let mut n = 0;
    for c in chunks {
        let mut ip = 0;
        while ip < c.code.len() {
            let op = Op::from_code_at(&c.code, ip);
            n += usize::from(op == Op::V128_LOAD);
            ip += 4;
            ip += op.operand_format().size_in(&c.code, ip);
        }
    }
    n
}
#[test]
fn parsed_c_storage_directive_order_aliases_and_byte_difference() {
    for fast in [false, true] {
        let chunks = compile(SOURCE, fast);
        let vectors = vector_count(&chunks);
        if fast {
            assert!(
                vectors >= 8,
                "Four physical memcmp sites must select real SIMD"
            );
        } else {
            assert_eq!(vectors, 0);
        }
        let native = vybe_platform_wasm::write_wasm(&chunks);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&native)
            .unwrap();
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        vm.run(chunks).unwrap();
        assert!(vm.last_exception.is_none());
        assert_eq!(vm.global("visits").unwrap().as_f64(), 12.0);
        assert!(vm.global("result").unwrap().as_f64() < 0.0);
    }
}

#[test]
fn direct_memcmp_comparisons_keep_the_declared_integer_result() {
    let source = r#"void *malloc(int size);
int memcmp(const void *a, const void *b, int size);
int main(void) {
    int *a = malloc(128), *b = malloc(128);
    for (int i = 0; i < 32; ++i) { a[i] = i + 17; b[i] = i + 17; }
    if (memcmp(a, b, 128) != 0) return 10;
    if (!(0 == memcmp(a, b, 128))) return 11;
    b[3] += 1;
    if (!(memcmp(a, b, 128) < 0)) return 12;
    if (!(memcmp(b, a, 128) > 0)) return 13;
    if (memcmp(a, b, 0) != 0) return 14;
    return 0;
}
"#;
    for fast in [false, true] {
        let chunks = compile(source, fast);
        for chunk in &chunks {
            assert!(
                !vybe_runtime::debug::disassemble(chunk).contains("dyneq"),
                "dynamic equality in {}",
                chunk.name
            );
        }
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm_with_options(
                &chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: false,
                },
            ))
            .unwrap();
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        vm.run(chunks).unwrap();
        assert_eq!(vm.pending_exit_code, 0);
    }
}
#[test]
fn user_memcmp_definition_retains_normal_call_resolution() {
    let source = r#"void *malloc(int n); int result; int memcmp(const void *a,const void *b,int n){return 77;} int main(){int *a=malloc(128);int *b=malloc(128);result=memcmp(a,b,128);return 0;}"#;
    let chunks = compile(source, true);
    assert_eq!(vector_count(&chunks), 0);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
        .unwrap();
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap();
    assert_eq!(vm.global("result").unwrap().as_f64(), 77.0);
}

#[test]
fn managed_buffers_and_internal_name_collision_retain_original_route() {
    for source in [
        r#"int result; int main(){char *a="same";char *b="same";result=memcmp(a,b,4);return 0;}"#,
        r#"void *malloc(int n);int result;int __c_ptr_byte_compare(void *a,void *b,int n){return 99;}int main(){int *a=malloc(128);int *b=malloc(128);result=memcmp(a,b,0);return 0;}"#,
    ] {
        let chunks = compile(source, true);
        assert_eq!(vector_count(&chunks), 0);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .unwrap();
        let mut vm = VM::new();
        primitives::platforms::register_platforms_all(&mut vm);
        vm.run(chunks).unwrap();
        assert_eq!(vm.global("result").unwrap().as_f64(), 0.0);
    }
}
