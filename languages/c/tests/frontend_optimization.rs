use std::path::Path;
use vybe_language_c::frontend::{Backend, with_backend};

fn parser_for(backend: Backend) -> vybe_runtime::registry::ParseFn {
    match backend {
        Backend::Peg => vybe_language_c::parse_with_peg,
        Backend::Native => vybe_language_c::parse_optimized,
        Backend::Owned => vybe_language_c::parse_owned_ast,
        Backend::OwnedContextual => vybe_language_c::parse_owned_contextual,
        Backend::ContextualPratt => vybe_language_c::parse_contextual_pratt,
    }
}

#[path = "support/frontend_unit_compatibility.rs"]
mod unit_compatibility;
use unit_compatibility::UNIT_BOUNDARY_PROGRAMS;

#[test]
fn committed_unit_boundaries_preserve_complete_typed_modules_and_rejections() {
    for source in UNIT_BOUNDARY_PROGRAMS.iter()
        .chain(unit_compatibility::USER_CALL_BINDING_PROGRAMS).copied().chain([
        "typedef struct Item Item; struct Item{int value;}; typedef Item Items[2]; \
         static Items rows[2]={{{1},{2}},{{3},{4}}}; int f(void){return rows[1][0].value+sizeof(Items);}",
        "typedef int T; static T a[3]; int f(int n,T b[n]){return sizeof(T[n])+b[n-1]+a[0];} \
         static T a[3]={1,2,3};",
        "typedef int (*F)(int,...); int f(int x,...){return x;} F callback=f; \
         int good(void){return callback(2,3)+1;} int broken(void){return 1+;}",
        "typedef int T; static T a[]={1,2}; int good(void){return a[0]+2;} int broken( {",
    ]) {
        let reference = vybe_language_c::parse_with_peg(source).map(|module| format!("{module:?}"));
        let ((actual, construction), matching) = vybe_language_c::frontend::profile_matching(||
            vybe_language_c::frontend::profile_construction(|| vybe_language_c::parse_optimized(source)));
        assert_eq!(actual.map(|module| format!("{module:?}")), reference, "{source}");
        assert!(matching.rules.iter().any(|row| row.rule == "frontend_unit"), "{matching:?}");
        assert_eq!(construction.declaration_handoffs, 0, "Native acquired owned declaration policy");
        if reference.is_ok() {
            assert!(construction.source_pratt_handoffs > 0, "{construction:?}: {source}");
        }
    }
}

#[test]
fn default_native_units_execute_identical_standard_wasm_without_dynamic_equality() {
    vybe_language_c::register();
    unit_compatibility::execute_identical_standard_wasm(vybe_language_c::profile_source(),
        [vybe_language_c::parse_with_peg, vybe_language_c::parse_optimized]);
}

fn report_measurement(line: String) {
    eprintln!("{line}");
    if let Some(path) = std::env::var_os("VYBE_C_FRONTEND_REPORT") {
        use std::io::Write;
        let mut file = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(path)
            .unwrap();
        writeln!(file, "{line}").unwrap();
    }
}

#[cfg(unix)]
fn cpu_ns() -> u128 {
    let mut time = libc::timespec {
        tv_sec: 0,
        tv_nsec: 0,
    };
    assert_eq!(
        unsafe { libc::clock_gettime(libc::CLOCK_THREAD_CPUTIME_ID, &mut time) },
        0
    );
    time.tv_sec as u128 * 1_000_000_000 + time.tv_nsec as u128
}

#[test]
fn phase_timing_preserves_complete_peg_ast() {
    const CHILD: &str = "VYBE_C_PHASE_TIMING_TEST_CHILD";
    if std::env::var_os(CHILD).is_some() {
        let module = vybe_language_c::parse_with_peg(
            "struct P { int x; }; int g=3; int f(struct P *p) { p->x += g; return p->x; }",
        )
        .unwrap();
        println!("phase_test_ast={module:?}");
        return;
    }
    let path = std::env::temp_dir().join(format!(
        "vybe-c-phase-timing-{}-{}.log",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    let run = |timed: bool| {
        let mut command = std::process::Command::new(std::env::current_exe().unwrap());
        command
            .args([
                "--exact",
                "phase_timing_preserves_complete_peg_ast",
                "--nocapture",
            ])
            .env(CHILD, "1")
            .env_remove("VYBE_C_DEBUG_TIMINGS")
            .env_remove("VYBE_C_DEBUG_WRITE_PREPROCESSED");
        if timed {
            command.env("VYBE_C_DEBUG_TIMINGS", &path);
        }
        let output = command.output().unwrap();
        assert!(
            output.status.success(),
            "{}",
            String::from_utf8_lossy(&output.stderr)
        );
        String::from_utf8(output.stdout)
            .unwrap()
            .lines()
            .find(|line| line.starts_with("phase_test_ast="))
            .unwrap()
            .to_owned()
    };
    let baseline = run(false);
    assert!(!path.exists());
    assert_eq!(baseline, run(true));
    let timings = std::fs::read_to_string(&path).unwrap();
    std::fs::remove_file(&path).unwrap();
    let phases: Vec<_> = timings
        .lines()
        .filter_map(|line| line.split_whitespace().next())
        .collect();
    let expected = [
        "support_generation",
        "normalize_entry",
        "normalize_type_guards",
        "normalize_linear_bytes_initial",
        "normalize_affine_bytes",
        "normalize_loop_backing",
        "normalize_word_provenance",
        "normalize_linear_bytes_after_backing",
        "normalize_storage_access",
        "normalize_static_helpers",
        "total",
    ];
    let start = phases
        .iter()
        .position(|phase| *phase == expected[0])
        .unwrap();
    assert_eq!(&phases[start..], expected);
}

#[test]
fn complete_ast_spans_directives_and_rejections_match_peg() {
    for source in [
        "int main(void) { return 1 + 2 * 3 << 1; }",
        "typedef unsigned long T; T f(T x) { T y=(T)3; { int T=2; y += T; } return x+y; }",
        "struct P { int x; int y; }; int f(struct P *p) { p->x += 3; return p->y++; }",
        "int f(int x) { return x; } int g(void) { int (*p)(int)=f; return p(3); }",
        "int f(int n, ...) { return n; } int main(void) { return f(0, 2, 3); }",
        "int a[]={0x1e-2,0xbe+2}; int main(void){return a[0]!=28;}",
        "double a[]={0x1.ep-2,1e+2}; int main(void){return a[1]!=100;}",
        "#define X 3\n#define F(x) ((x)+X)\nint f(void) { return F(2); }",
        "int f(void) { int a[3]={[1]=2}; return a[1]; }",
        "int f(int x) { for(int i=0;i<3;i++) x+=i; return x ? x : 2; }",
        "int f(void) { return 0xffffffffu > 0; }",
        "long f(void) { return 0x100000000LL; }",
        "unsigned long f(void) { return 0xffffffffffffffffULL >> 1; }",
        "int f(void) { return 0xffffffffu/2u + 0x80000000%3u; }",
        "int f(void) { return 1+2*3==7!=0; }",
        "#define X 2\nlong f(void) { return 1ULL+X*3; }",
        "int f(int x) { return 1+2*3+x; }",
        "int f(int x) { return 1==2||x; }",
        "long f(int x) { return 1<<40|x; }",
        "int f(void) { return 1+2.0; }",
        "int f(void) { return 1+-2; }",
        "int f(void) { return 1+(2+3); }",
        "float f(void) { return -0x1.8p+2 + .5f; }",
        "int f(void) { return -~!+2 + L'\\u00ff' + u8'\\x41'; }",
        "int f(void) { static unsigned a[]={-1u,~2u,'a',!0}; return a[0]; }",
        "typedef unsigned T; int f(T x) { { int x=3; x += -2; } return !x; }",
        "int f(int *p) { return *p + sizeof(*p) + offsetof(struct X, y); }",
        "int f(void) { return -; }",
        "int f(void) { return 0x1.p+; }",
        "int f(void) { return 1 /* operator */ + 2 /* tail */; }",
        "int f(void) { return sizeof(1==2) + _Alignof(int); }",
        "int f(void) { return 1 + sizeof(int) + 1+(int)2; }",
        "int f(void) { return 1 + _Generic(2,int:3,default:4); }",
        "int f(void) { int a[1+2]={1,2,3}; return a[1]; }",
        "int f(void) { return \"a\\n\" \"b\"[0]; }",
        "int f(void) { return 1 + ; }",
        "int f(void) { return 1+2*3; } @",
        "int f(void) { return \"unterminated; }",
        "typedef int T; int f(void) { return (T; }",
    ] {
        let reference = vybe_language_c::parse_with_peg(source).map(|module| format!("{module:?}"));
        let candidate =
            vybe_language_c::parse_optimized(source).map(|module| format!("{module:?}"));
        assert_eq!(reference, candidate, "{source}");
        assert_eq!(
            reference,
            vybe_language_c::parse_owned_ast(source).map(|module| format!("{module:?}")),
            "owned: {source}"
        );
        assert_eq!(
            reference,
            with_backend(Backend::OwnedContextual, || vybe_language_c::parse(source))
                .map(|module| format!("{module:?}")),
            "mixed: {source}"
        );
        assert_eq!(
            reference,
            vybe_language_c::parse_contextual_pratt(source).map(|module| format!("{module:?}")),
            "contextual Pratt: {source}"
        );
    }
}

#[test]
fn production_owned_declarations_keep_ordered_scope_and_initializer_policy() {
    for source in [
        "int x=2; int y=3;",
        "int main(void){int x=2,y=x+3; return y!=5;}",
        "typedef unsigned word; static word x=1u; int f(void){word x=2u; return x;} int main(void){return f()!=2;}",
        "int x=1; int main(void){int x=2; {int x=3;} static int y=4; const int z=5; return x+y+z!=11;}",
        "unsigned char x=257; unsigned short y=65537; unsigned z=0xffffffffu; double f=1.25; int c=L'\\n';",
        "typedef int V[2]; V x; typedef int *P; P y=0; int main(void){return y!=0;}",
        "int a[2]={1,2}; int x=3; int main(void){int y=4; return a[0]+x+y!=8;}",
    ] {
        let (reference, baseline) = vybe_language_c::frontend::profile_construction(|| {
            vybe_language_c::parse_with_peg(source)
        });
        let (candidate, profile) = vybe_language_c::frontend::profile_construction(|| {
            vybe_language_c::parse_owned_contextual(source)
        });
        assert_eq!(
            format!("{:?}", reference.unwrap()),
            format!("{:?}", candidate.unwrap()),
            "{source}"
        );
        assert!(
            profile.declaration_handoffs > 0,
            "no production declaration handoff: {source}"
        );
        assert_eq!(
            profile.parse_calls, baseline.parse_calls,
            "extra parse: {source}"
        );
    }
}

#[test]
fn production_native_initializers_consume_owned_ast_and_keep_contextual_fallback() {
    for source in [
        "int a[2][3]={{1,2,3},{4,5,6}}; int main(void){return a[1][2]!=6;}",
        "struct S{int x;int y;}; struct S a[]={{1,2},{3,4}};",
        "unsigned a[]={1u,0xffffffffu}; double b[]={0.5,0x1p1};",
        "int f(void){int x=1; int a[]={x,2}; int b[]={3,4}; return a[0]+b[1];}",
        "int f(void){int a[3]={[1]=2}; int b[]={3,4}; return a[1]+b[0];}",
    ] {
        let reference = vybe_language_c::parse_with_peg(source).unwrap();
        let (candidate, counts) = vybe_language_c::frontend::profile_construction(|| {
            vybe_language_c::parse_optimized(source).unwrap()
        });
        assert_eq!(format!("{reference:?}"), format!("{candidate:?}"), "{source}");
        assert!(counts.direct_initializer_handoffs > 0,
            "production native parsing did not consume direct AST: {counts:?}: {source}");
    }
}

#[test]
fn staged_scalar_declarations_move_ast_with_explicit_scope_policies() {
    use vybe_language_c::frontend::{
        DeclarationContext, DeclarationPolicy, ScalarPolicy, UnsignedWidth,
        construct_scalar_declaration,
    };
    for (source, integer_wrap) in [
        ("int x=2;", None),
        ("int x=2,y;", None),
        ("unsigned x=1u;", Some(UnsignedWidth::W32)),
        ("double x=1.25;", None),
        ("double x=0x1.8p+2;", None),
        ("int x='a';", None),
        ("int x=L'\\n';", None),
        ("int x=u8'a';", None),
        ("int x=true;", None),
        ("int x=-2;", None),
        ("int x=~2;", None),
        ("int x=!2;", None),
        ("int x=2+3*4;", None),
        ("int x=1==2;", None),
        ("int x=1<<40;", None),
    ] {
        let reference = vybe_language_c::parse_with_peg(source).unwrap();
        let statement = reference
            .body
            .iter()
            .find(|statement| {
                matches!(&statement.kind, vybe_ast::StmtKind::VarDecl { declarations, .. }
                if declarations.iter().any(|decl| matches!(&decl.pattern,
                    vybe_ast::BindingPattern::Ident(name) if name == "x")))
            })
            .unwrap();
        let vybe_ast::StmtKind::VarDecl { declarations, kind } = &statement.kind else {
            unreachable!()
        };
        let mut context = DeclarationContext::new(source);
        context.declarations.insert(
            (0, source.len()),
            DeclarationPolicy {
                kind: kind.clone(),
                statement_span: statement.span,
            },
        );
        let start = source.find('x').unwrap();
        let end = source[start..].find([',', ';']).unwrap() + start;
        context.declarators.insert(
            (start, end),
            ScalarPolicy {
                emitted_name: "x".into(),
                type_hint: declarations[0].type_hint.clone(),
                zero_if_missing: declarations[0].init.is_some(),
                integer_wrap,
            },
        );
        if let Some(start) = source.find('y') {
            context.declarators.insert(
                (start, start + 1),
                ScalarPolicy {
                    emitted_name: "y".into(),
                    type_hint: declarations[1].type_hint.clone(),
                    zero_if_missing: declarations[1].init.is_some(),
                    integer_wrap,
                },
            );
        }
        let actual = construct_scalar_declaration(source, context, grameo::MatchOptions::default())
            .unwrap()
            .unwrap();
        assert_eq!(format!("{statement:?}"), format!("{actual:?}"), "{source}");
    }
}

#[test]
fn expanded_real_sources_keep_include_and_macro_order() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    for name in [
        "examples/c/sdlcalc/calc.c",
        "examples/c/doom_funcptr_probe.c",
        "tests/c/doom_startup/multi_glob_conditions.c",
    ] {
        let path = root.join(name);
        let source = std::fs::read_to_string(&path).unwrap();
        let expanded = vybe_language_c::walker::expand_source(&path, &source).unwrap();
        let reference = vybe_language_c::parse_with_peg(&expanded).unwrap();
        let candidate = vybe_language_c::parse_optimized(&expanded).unwrap();
        assert_eq!(format!("{reference:?}"), format!("{candidate:?}"), "{name}");
        let owned = vybe_language_c::parse_owned_ast(&expanded).unwrap();
        assert_eq!(
            format!("{reference:?}"),
            format!("{owned:?}"),
            "owned: {name}"
        );
        let pratt = vybe_language_c::parse_contextual_pratt(&expanded).unwrap();
        assert_eq!(
            format!("{reference:?}"),
            format!("{pratt:?}"),
            "contextual Pratt: {name}"
        );
    }
}

#[test]
fn environment_selected_production_dispatch_matches_peg_in_isolated_processes() {
    const CHILD: &str = "VYBE_C_PRODUCTION_DISPATCH_TEST_CHILD";
    const TEST: &str =
        "environment_selected_production_dispatch_matches_peg_in_isolated_processes";
    if std::env::var_os(CHILD).is_some() {
        let selection = std::env::var("VYBE_C_FRONTEND").unwrap();
        let expected_parser: vybe_runtime::registry::ParseFn = match selection.as_str() {
            "peg" => vybe_language_c::parse_with_peg,
            "pratt" => vybe_language_c::parse_contextual_pratt,
            other => panic!("unexpected child backend: {other}"),
        };
        vybe_language_c::register();
        let selected = vybe_runtime::registry::find("c").unwrap().parse;
        assert!(std::ptr::fn_addr_eq(
            selected,
            expected_parser,
        ));
        for source in [
            "typedef unsigned T; int f(T x) { { int x=3; x += -2; } return !x + 2*3; }",
            "struct P { int x; }; int f(struct P *p,int *q) { p->x += (*q)++; return p->x + *q*2; }",
            "int f(void) { return 1 + ; }",
            "int f(void) { return 1+2*3; } @",
        ] {
            let expected = vybe_language_c::parse_with_peg(source)
                .map(|module| format!("{module:?}"));
            // A new thread must select the registered backend without inheriting TLS.
            let (actual, counts) = std::thread::spawn(move || {
                vybe_language_c::frontend::profile_construction(|| {
                    selected(source).map(|module| format!("{module:?}"))
                })
            })
            .join()
            .unwrap();
            assert_eq!(expected, actual, "{selection}: {source}");
            if actual.is_ok() && selection == "pratt" {
                assert!(counts.contextual_pratt_handoffs > 0);
            } else {
                assert_eq!(counts.contextual_pratt_handoffs, 0);
            }
        }
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        for name in [
            "examples/c/sdlcalc/calc.c",
            "examples/c/doom_funcptr_probe.c",
            "tests/c/doom_startup/multi_glob_conditions.c",
        ] {
            let path = root.join(name);
            let mut reference = vybe_compiler::projects::load(&path).unwrap();
            reference.language.parse = vybe_language_c::parse_with_peg;
            let expected = reference.prepared_module().unwrap();
            let production = vybe_compiler::projects::load(&path).unwrap();
            assert!(std::ptr::fn_addr_eq(production.language.parse, selected));
            let (actual, counts) = vybe_language_c::frontend::profile_construction(|| {
                production.prepared_module().unwrap()
            });
            assert_eq!(
                format!("{expected:?}"),
                format!("{actual:?}"),
                "production {selection}: {name}"
            );
            if selection == "pratt" {
                assert!(counts.contextual_pratt_handoffs > 0, "{name}");
            } else {
                assert_eq!(counts.contextual_pratt_handoffs, 0, "{name}");
            }
        }
        println!("production_dispatch_parity={selection}");
        return;
    }
    for selection in ["peg", "pratt"] {
        let output = std::process::Command::new(std::env::current_exe().unwrap())
            .args(["--exact", TEST, "--nocapture", "--test-threads=1"])
            .env(CHILD, "1")
            .env("VYBE_C_FRONTEND", selection)
            .env_remove("VYBE_C_DEBUG_TIMINGS")
            .env_remove("VYBE_C_DEBUG_WRITE_PREPROCESSED")
            .env_remove("VYBE_BUNDLE_DEBUG_TIMINGS")
            .env_remove("VYBE_COMPILER_DEBUG_TIMINGS")
            .output()
            .unwrap();
        assert!(
            output.status.success(),
            "{selection}: {:?}\n{}\n{}",
            output.status.code(),
            String::from_utf8_lossy(&output.stdout),
            String::from_utf8_lossy(&output.stderr)
        );
        assert!(
            String::from_utf8_lossy(&output.stdout)
                .contains(&format!("production_dispatch_parity={selection}"))
        );
    }
}

#[test]
fn public_mixed_entry_uses_contextual_handoff_and_keeps_complete_ast() {
    let source = "struct P { int x; }; int f(struct P *p,int x) { { int x=3; p->x += x; } return p->x + x*2; }";
    let reference = vybe_language_c::parse_with_peg(source).unwrap();
    let (actual, profile) = vybe_language_c::frontend::profile_construction(|| {
        vybe_language_c::parse_owned_contextual(source)
    });
    assert_eq!(format!("{reference:?}"), format!("{:?}", actual.unwrap()));
    assert!(profile.mixed_regions > 0);
    assert!(profile.mixed_handoffs > 0);
    assert!(profile.contextual_operands > 0);
}

#[test]
#[ignore = "bounded real-source construction and phase attribution"]
fn profile_complete_frontend() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let name = std::env::var("VYBE_C_FRONTEND_SOURCE")
        .unwrap_or_else(|_| "examples/chocolate-doom/src/doom/info.c".into());
    let path = root.join(&name);
    let source = std::fs::read_to_string(&path).unwrap();
    let expanded = vybe_language_c::walker::expand_source(&path, &source).unwrap();
    for backend in [Backend::Peg, Backend::Owned, Backend::OwnedContextual, Backend::ContextualPratt] {
        let (module, counts) = vybe_language_c::frontend::profile_construction(|| {
            with_backend(backend, || vybe_language_c::parse(&expanded))
        });
        let module = module.unwrap();
        eprintln!(
            "construction {name} {backend:?} {counts:?} statements={}",
            module.body.len()
        );
        std::hint::black_box(module);
    }
}

#[test]
fn expected_execution_and_standard_wasm_match() {
    vybe_language_c::register();
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    for name in [
        "tests/c/doom_startup/static_wide_comparisons.c",
        "tests/c/doom_startup/native_word_cast.c",
        "tests/c/doom_startup/bounded_varargs_format.c",
        "tests/c/doom_startup/explicit_main_arguments.c",
        "tests/c/doom_startup/multi_glob_conditions.c",
        "tests/c/doom_startup/frontend_initializer_handoffs.c",
        "tests/c/doom_startup/frontend_numeric_prefix_handoffs.c",
    ] {
        let path = root.join(name);
        let mut output = Vec::new();
        for backend in [
            Backend::Peg,
            Backend::Native,
            Backend::Owned,
            Backend::OwnedContextual,
            Backend::ContextualPratt,
        ] {
            let mut bundle = vybe_compiler::projects::load(&path).unwrap();
            bundle.language.parse = parser_for(backend);
            let chunks = bundle.compile().unwrap();
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: true,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap();
            let mut vm = vybe_runtime::VM::new();
            vybe_runtime::init_all_registered(
                &mut vm,
                &vybe_runtime::capabilities::Capabilities::all(),
            );
            vm.run(chunks).unwrap();
            assert_eq!(vm.pending_exit_code, 0, "{name} {backend:?}");
            output.push(bytes);
        }
        assert_eq!(output[0], output[1], "standard Wasm differs: {name}");
        assert_eq!(output[0], output[2], "owned standard Wasm differs: {name}");
        assert_eq!(output[0], output[3], "mixed standard Wasm differs: {name}");
        assert_eq!(output[0], output[4], "Pratt standard Wasm differs: {name}");
    }
}

#[test]
#[cfg(unix)]
#[ignore = "paired ordinary-debug whole-frontend measurement"]
fn measure_complete_frontend() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let names = std::env::var("VYBE_C_FRONTEND_SOURCE")
        .map(|name| vec![name])
        .unwrap_or_else(|_| {
            vec![
                "examples/c/sdlcalc/calc.c".into(),
                "examples/c/doom_funcptr_probe.c".into(),
                "tests/c/doom_startup/multi_glob_conditions.c".into(),
            ]
        });
    let samples = std::env::var("VYBE_C_FRONTEND_SAMPLES")
        .ok()
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(10);
    assert!(samples >= 3);
    for name in names {
        let path = root.join(&name);
        let source = std::fs::read_to_string(&path).unwrap();
        let expanded = vybe_language_c::walker::expand_source(&path, &source).unwrap();
        let backends = [
            Backend::Peg,
            Backend::Native,
            Backend::Owned,
            Backend::OwnedContextual,
            Backend::ContextualPratt,
        ];
        let mut times = [Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new()];
        let mut cpus = [Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new()];
        for sample in 0..samples + 2 {
            for index in if sample % 2 == 0 {
                [0, 1, 2, 3, 4]
            } else {
                [4, 3, 2, 1, 0]
            } {
                let start = std::time::Instant::now();
                let cpu = cpu_ns();
                let module =
                    with_backend(backends[index], || vybe_language_c::parse(&expanded)).unwrap();
                std::hint::black_box(module);
                if sample >= 2 {
                    cpus[index].push(cpu_ns() - cpu);
                    times[index].push(start.elapsed().as_nanos());
                }
            }
        }
        for (index, values) in times.iter_mut().enumerate() {
            values.sort_unstable();
            cpus[index].sort_unstable();
            report_measurement(format!(
                "frontend {name} {:?} wall_median_ns={} cpu_median_ns={} samples={}",
                backends[index],
                values[values.len() / 2],
                cpus[index][values.len() / 2],
                values.len()
            ));
        }
    }
}

#[test]
#[cfg(unix)]
#[ignore = "paired debug source-load, preparation, parse/walk and complete compilation"]
fn measure_complete_compilation() {
    vybe_language_c::register();
    let mut vm = vybe_runtime::VM::new();
    vybe_runtime::init_all_registered(&mut vm, &vybe_runtime::capabilities::Capabilities::all());
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let name = std::env::var("VYBE_C_FRONTEND_SOURCE")
        .unwrap_or_else(|_| "examples/chocolate-doom/src/doom/m_random.c".into());
    let path = root.join(&name);
    let samples = std::env::var("VYBE_C_FRONTEND_SAMPLES")
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(6);
    assert!(samples >= 3);
    let backends = [Backend::Peg, Backend::Native, Backend::Owned,
        Backend::OwnedContextual, Backend::ContextualPratt];
    let mut reference = None;
    for backend in backends {
        let mut bundle = vybe_compiler::projects::load(&path).unwrap();
        bundle.language.parse = parser_for(backend);
        let module = format!("{:?}", bundle.prepared_module().unwrap());
        if let Some(reference) = &reference {
            assert_eq!(reference, &module, "complete preparation differs: {backend:?}");
        } else {
            reference = Some(module);
        }
    }
    let mut wall = [Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new()];
    let mut cpu = [Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new()];
    let process_cpu = || {
        let mut time = libc::timespec {
            tv_sec: 0,
            tv_nsec: 0,
        };
        assert_eq!(
            unsafe { libc::clock_gettime(libc::CLOCK_PROCESS_CPUTIME_ID, &mut time) },
            0
        );
        time.tv_sec as u128 * 1_000_000_000 + time.tv_nsec as u128
    };
    for sample in 0..samples + 2 {
        for index in if sample % 2 == 0 {
            [0, 1, 2, 3, 4]
        } else {
            [4, 3, 2, 1, 0]
        } {
            let started = std::time::Instant::now();
            let before = process_cpu();
            let mut bundle = vybe_compiler::projects::load(&path).unwrap();
            bundle.language.parse = parser_for(backends[index]);
            let chunks = bundle.compile().unwrap();
            let elapsed = started.elapsed().as_nanos();
            let used = process_cpu() - before;
            assert!(!chunks.is_empty());
            std::hint::black_box(chunks);
            if sample >= 2 {
                wall[index].push(elapsed);
                cpu[index].push(used);
            }
        }
    }
    for (index, backend) in backends.into_iter().enumerate() {
        wall[index].sort_unstable();
        cpu[index].sort_unstable();
        report_measurement(format!(
            "compilation {name} {backend:?} wall_median_ns={} process_cpu_median_ns={} samples={samples}",
            wall[index][samples / 2],
            cpu[index][samples / 2]
        ));
    }
}

#[test]
#[ignore = "single expanded actual-Doom frontend phase profile"]
fn profile_expanded_doom_frontend() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/chocolate-doom/vybe_doom_core.c");
    let source = std::fs::read_to_string(&path).unwrap();
    let started = std::time::Instant::now();
    let expanded = vybe_language_c::walker::expand_source(&path, &source).unwrap();
    eprintln!(
        "profile backend=Peg input={} expanded_bytes={} expansion={:?}",
        path.display(),
        expanded.len(),
        started.elapsed()
    );
    let started = std::time::Instant::now();
    let module = with_backend(Backend::Peg, || vybe_language_c::parse(&expanded)).unwrap();
    eprintln!(
        "profile backend=Peg frontend={:?} module_statements={}",
        started.elapsed(),
        module.body.len()
    );
    std::hint::black_box(module);
}
