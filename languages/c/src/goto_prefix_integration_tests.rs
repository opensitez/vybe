use crate::frontend::{Backend, with_backend};

fn candidate(source: &str, backend: Backend) -> Result<vybe_ast::Module, String> {
    with_backend(backend, || crate::walker::parse_with_function_goto_policy(source, true))
}

fn peg(source: &str) -> Result<vybe_ast::Module, String> { candidate(source, Backend::Peg) }
fn native(source: &str) -> Result<vybe_ast::Module, String> { candidate(source, Backend::Native) }

const PROGRAMS: &[&str] = &[
    include_str!("../../../tests/c/doom_startup/nested_label_goto.c"),
    "static int calls;int init(void){calls++;return 0;}int main(void){int n=init();again:n++;if(n<3)goto again;return calls!=1 || n!=3;}",
    "int main(void){int n=0;start:n++;int local=n;if(n<3)goto start;return n!=3 || local!=3;}",
    "int route(int enter){int n=0;if(enter)goto body;switch(0){default:n+=2;body:n+=3;case 1:n+=4;}return n;}int main(void){return route(1)!=7 || route(0)!=9;}",
    "int main(void){int value=3;int *p=&value;goto inside;if(0){inside:*p+=4;}return value!=7;}",
    "struct S{int n;};int main(void){struct S s={3};goto inside;if(0){inside:s.n+=4;}return s.n!=7;}",
    "int f(void){int n=0;goto inside;if(0){inside:n=3;}return n;}int g(void){int n=0;goto inside;if(0){inside:n=4;}return n;}int main(void){return f()!=3 || g()!=4;}",
    "int main(void){int __c_function_goto_pc_0=2;int n=0;goto inside;if(0){inside:n=__c_function_goto_pc_0;}return n!=2;}",
];

#[test]
fn prefix_function_hook_preserves_complete_backend_ast() {
    for source in PROGRAMS {
        let expected = format!("{:?}", peg(source).unwrap());
        assert_eq!(format!("{:?}", native(source).unwrap()), expected, "{source}");
        assert!(!expected.contains("GoTo("), "{source}");
    }
}

#[test]
fn prefix_function_hook_rejects_without_leaking_parse_scope() {
    for backend in [Backend::Peg, Backend::Native] {
        for source in [
            "int main(void){goto inside;{int local=3;inside:return 0;}}",
            "int main(void){goto top;int late=3;top:goto inside;if(0){inside:return late;}}",
            "int main(void){goto missing;return 0;}",
        ] {
            let error = candidate(source, backend).unwrap_err();
            assert!(error.contains("C goto preparation in main"), "{error}");
        }
        let original = with_backend(backend, || crate::walker::parse(PROGRAMS[0])).unwrap();
        assert!(format!("{original:?}").contains("GoTo("));
        assert!(candidate(PROGRAMS[0], backend).is_ok());
    }
}

#[test]
fn prefix_function_hook_leaves_no_goto_functions_identical() {
    for source in [
        "int add(int a,int b){return a+b;}int main(void){return add(2,3)!=5;}",
        "int main(void){int n=1;{int n=3;n++;}return n!=1;}",
        "int main(void){int n=0;for(int i=0;i<3;i++){n+=i;}return n!=3;}",
    ] {
        for backend in [Backend::Peg, Backend::Native] {
            let reference = with_backend(backend, || crate::walker::parse(source)).unwrap();
            assert_eq!(format!("{:?}", candidate(source, backend).unwrap()), format!("{reference:?}"));
        }
    }
}

#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
fn prefix_function_hook_executes_standard_wasm_without_dynamic_equality() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(
        crate::profile_source(), [peg, native], PROGRAMS.iter().copied());
}

#[test]
#[ignore = "explicit saved-source whole-function admission audit"]
fn saved_doom_reports_every_prefix_policy_decline() {
    let path = std::env::var("NYX_C_GOTO_DOOM_SOURCE").expect("set saved preprocessed C source path");
    let source = std::fs::read_to_string(path).unwrap();
    let module = crate::walker::with_unlowered_gotos_for_tests(|| crate::parse_optimized(&source)).unwrap();
    let mut counts = [0usize; 4];
    for statement in &module.body {
        if let vybe_ast::StmtKind::FunctionDecl { name, body, params, .. } = &statement.kind {
            let names = params.iter().map(|parameter| parameter.name.clone()).collect();
            match crate::goto_prefix_policy::prepare(body, &names) {
                Ok(prepared) => counts[match prepared.path {
                    crate::goto_prefix_policy::Path::Unchanged => 0,
                    crate::goto_prefix_policy::Path::PrefixGraph => 1,
                    crate::goto_prefix_policy::Path::ProvenTopTargets => 2,
                }] += 1,
                Err(error) => {
                    counts[3] += 1;
                    eprintln!("declined function {name}: {} at {:?}", error.reason, error.span);
                }
            }
        }
    }
    eprintln!("unchanged/graph/top/declined={counts:?}; source_bytes={}", source.len());
    assert_eq!(counts[3], 0, "all explicit declines must be reviewed before promotion");
}
