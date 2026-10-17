use vybe_compiler::{Compiler, primitives};
use vybe_runtime::{Chunk, Op, VM, profile::BuiltinEmit};

const SOURCE: &str = r#"
void *malloc(int n);
int strcmp(const char *a, const char *b);
int strncmp(const char *a, const char *b, int n);
int result; int visits; int failures;
int main() {
  int *a=malloc(128); int *b=malloc(128);
  for(int i=0;i<32;i++){a[i]=16843009;b[i]=16843009;}
  a[24]=0; b[24]=0;
  result=strcmp((char*)(visits=visits*10+1,a),(char*)(visits=visits*10+2,b));
  if(result!=0 || visits!=12) failures=failures+1;
  result=strncmp((char*)a,(char*)b,0); if(result!=0) failures=failures+1;
  result=strcmp((char*)a,(char*)a); if(result!=0) failures=failures+1;
  b[25]=16843010;
  result=strcmp((char*)a,(char*)b); if(result!=0) failures=failures+1;
  b[4]=16843010;
  result=strncmp((char*)a,(char*)b,16); if(result!=0) failures=failures+1;
  result=strncmp((char*)a,(char*)b,17); if(result>=0) failures=failures+1;
  result=strcmp((char*)a,(char*)b); if(result>=0) failures=failures+1;
  return failures;
}
"#;

fn compile(source: &str, fast: bool) -> Vec<Chunk> {
    vybe_language_c::register();
    let ast = vybe_language_c::parse(source).unwrap();
    let mut profile =
        vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap();
    if !fast {
        for (name, emitter) in [
            ("__c_ptr_string_compare", "c.strcmp"),
            ("__c_ptr_string_ncompare", "c.strncmp"),
        ] {
            profile.builtins.get_mut(name).unwrap().emit = BuiltinEmit::Common(emitter.into());
        }
    }
    Compiler::with_profile(profile).compile(&ast).unwrap()
}

fn vectors(chunks: &[Chunk]) -> usize {
    let mut count = 0;
    for c in chunks {
        let mut ip = 0;
        while ip < c.code.len() {
            let op = Op::from_code_at(&c.code, ip);
            count += usize::from(op == Op::V128_LOAD);
            ip += 4;
            ip += op.operand_format().size_in(&c.code, ip);
        }
    }
    count
}

fn run(chunks: Vec<Chunk>) -> VM {
    let native = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&native)
        .unwrap();
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap();
    assert!(vm.last_exception.is_none());
    vm
}

#[test]
fn parsed_c_linear_strings_preserve_nul_bounds_alias_and_argument_order() {
    for fast in [false, true] {
        let chunks = compile(SOURCE, fast);
        let vector_count = vectors(&chunks);
        if fast {
            assert!(vector_count >= 21);
        } else {
            assert_eq!(vector_count, 0);
        }
        let vm = run(chunks);
        assert_eq!(vm.global("failures").unwrap().as_f64(), 0.0);
        assert_eq!(vm.global("visits").unwrap().as_f64(), 12.0);
        assert!(vm.global("result").unwrap().as_f64() < 0.0);
    }
}

#[test]
fn user_definitions_managed_storage_and_internal_collisions_keep_their_routes() {
    for (source, expected) in [
        (
            r#"void *malloc(int n);int result;int strcmp(const char *a,const char *b){return 77;}int main(){int *a=malloc(128);int *b=malloc(128);result=strcmp((char*)a,(char*)b);return 0;}"#,
            77.0,
        ),
        (
            r#"void *malloc(int n);int result;int strncmp(const char *a,const char *b,int n){return 78;}int main(){int *a=malloc(128);int *b=malloc(128);result=strncmp((char*)a,(char*)b,0);return 0;}"#,
            78.0,
        ),
        (
            r#"int result;int main(){char *a="same";char *b="same";result=strcmp(a,b);return 0;}"#,
            0.0,
        ),
        (
            r#"void *malloc(int n);int result;int __c_ptr_string_compare(char *a,char *b){return 99;}int main(){int *a=malloc(128);int *b=malloc(128);a[0]=0;b[0]=0;result=strcmp((char*)a,(char*)b);return 0;}"#,
            0.0,
        ),
        (
            r#"void *malloc(int n);int result;int __c_ptr_string_ncompare(char *a,char *b,int n){return 99;}int main(){int *a=malloc(128);int *b=malloc(128);result=strncmp((char*)a,(char*)b,0);return 0;}"#,
            0.0,
        ),
    ] {
        let chunks = compile(source, true);
        assert_eq!(vectors(&chunks), 0);
        assert_eq!(run(chunks).global("result").unwrap().as_f64(), expected);
    }
}
