use super::*;

const PROGRAMS: &[&str] = &[
    "static int first[]={1,2}; int f(int x){return first[x]+x*2;} \
     static int second[]={3,4}; int g(int x){return second[x]+f(x);} \
     int main(void){return f(1)+g(0)!=8;}",
    "typedef unsigned short T; int total(int n){T sum=1; int a[n]; \
     for(int i=0;i<n;i++){a[i]=i+1;sum+=a[i];} \
     {int T=7;sum+=T;} return sum;} int main(void){return total(3)!=14;}",
    "typedef struct Node Node; struct Node{int value;Node *next;}; extern Node *head; \
     int read(void){return head->value+head[1].value;} \
     static Node nodes[2]={{3,0},{4,0}}; Node *head=nodes; \
     int main(void){return read()!=7;}",
    "typedef int (*Callback)(int,...); int choose(int x,...){return x;} \
     static int values[]={3,4}; static Callback selected=choose; \
     int invoke(Callback cb){return cb(values[0],values[1])+values[1];} \
     int main(void){return invoke(selected)!=7;}",
    "typedef int (*Callback)(int,...); int choose(int x,...){return x;} \
     static int matrix[2][2]={{1,2},{3,4}}; \
     static int (*rows[2])[2]={&matrix[0],&matrix[1]}; \
     static Callback callbacks[2]={choose,choose}; \
     int main(void){return callbacks[1]((*rows[1])[1],3)!=4;}",
    "static int result; void step(int n){int i=0;while(i<n){i++; \
     if(i==2)continue;if(i==5)break;result+=i;} \
     if(result==8)goto done;result=-1;done:return;} \
     int main(void){step(8);return result!=8;}",
    "typedef unsigned short (*Callback)(void); extern Callback callbacks[2]; \
     int invoke(void){return callbacks[1]()!=60001;} \
     unsigned short choose(void){return 60001;} Callback callbacks[2]={choose,choose}; \
     int main(void){return invoke();}",
    "typedef int (*Callback)(int); int choose(int x){return x;} Callback selected=choose; \
     Callback *slots[1]={&selected}; int main(void){return (*slots[0])(7)!=7;}",
];

const WINDOWED_CALLBACK: &str = "#include <SDL.h>\n\
    int worker(int x){return x+3;} \
    int main(void){if(0)SDL_CreateWindow(\"unit\",0,0,1,1,0);return worker(0)-3;}";

fn with_unit_engine<T>(general: bool, action: impl FnOnce() -> T) -> T {
    struct Restore {
        general: bool,
        native_pratt: Option<bool>,
        pratt_disabled: bool,
        direct_disabled: bool,
    }
    impl Drop for Restore {
        fn drop(&mut self) {
            GENERAL_UNIT_ENGINE.with(|slot| slot.set(self.general));
            NATIVE_PRATT_ENABLED.with(|slot| slot.set(self.native_pratt));
            SOURCE_PRATT_DISABLED.with(|slot| slot.set(self.pratt_disabled));
            DIRECT_DISABLED.with(|slot| slot.set(self.direct_disabled));
        }
    }
    let _restore = Restore {
        general: GENERAL_UNIT_ENGINE.with(|slot| slot.replace(general)),
        native_pratt: NATIVE_PRATT_ENABLED.with(|slot| slot.replace(Some(true))),
        pratt_disabled: SOURCE_PRATT_DISABLED.with(|slot| slot.replace(false)),
        direct_disabled: DIRECT_DISABLED.with(|slot| slot.replace(false)),
    };
    action()
}

fn parse_generated(source: &str) -> Result<vybe_ast::Module, String> {
    with_unit_engine(false, || crate::parse_optimized(source))
}

fn parse_general(source: &str) -> Result<vybe_ast::Module, String> {
    with_unit_engine(true, || crate::parse_optimized(source))
}

fn assert_engine_counts(profile: &ConstructionProfile, general: bool) {
    if general {
        assert!(profile.general_unit_matches > 0, "{profile:?}");
        assert_eq!(profile.generated_unit_matches, 0, "{profile:?}");
    } else {
        assert!(profile.generated_unit_matches > 0, "{profile:?}");
        assert_eq!(profile.general_unit_matches, 0, "{profile:?}");
    }
}

#[test]
fn generated_units_preserve_complete_typed_modules_and_actual_handoffs() {
    let mut pratt = [0; 2];
    let mut initializers = [0; 2];
    let mut bodies = [0; 2];
    let mut primaries = [0; 2];
    for source in PROGRAMS.iter().copied().chain([WINDOWED_CALLBACK]) {
        let expected = crate::parse_with_peg(source).unwrap();
        if source == WINDOWED_CALLBACK {
            let mut callbacks = 0;
            for statement in &expected.body {
                statement.walk_exprs(&mut |value| {
                    callbacks +=
                        usize::from(matches!(value.kind, vybe_ast::ExprKind::FunctionExpr(_)));
                });
            }
            assert!(
                callbacks > 0,
                "fixture must retain the actual SDL guest FunctionExpr"
            );
        }
        for general in [false, true] {
            let ((actual, construction), matching) = with_unit_engine(general, || {
                profile_matching(|| profile_construction(|| crate::parse_optimized(source)))
            });
            assert_eq!(
                format!("{:?}", actual.unwrap()),
                format!("{expected:?}"),
                "general={general}: {source}"
            );
            assert_engine_counts(&construction, general);
            assert!(construction.source_pratt_matches > 1, "{construction:?}");
            assert_eq!(construction.source_pratt_recaptures, 0, "{construction:?}");
            assert!(
                matching
                    .rules
                    .iter()
                    .any(|row| row.rule == Rule::frontend_unit.name() && row.totals.calls > 1),
                "{matching:?}"
            );
            assert!(
                !matching
                    .rules
                    .iter()
                    .any(|row| row.rule == Rule::program.name()),
                "valid units must not be whole-program recaptured: {matching:?}"
            );
            let index = usize::from(general);
            pratt[index] += construction.source_pratt_handoffs;
            initializers[index] += construction.direct_initializer_handoffs;
            bodies[index] += construction.direct_body_handoffs;
            primaries[index] += construction.direct_primary_handoffs;
        }
    }
    for index in 0..2 {
        assert!(
            pratt[index] > 0,
            "no typed Pratt handoffs, general={}",
            index == 1
        );
        assert!(
            initializers[index] > 0,
            "no Direct initializer handoffs, general={}",
            index == 1
        );
        assert!(
            bodies[index] > 0,
            "no Direct body handoffs, general={}",
            index == 1
        );
        assert!(
            primaries[index] > 0,
            "no Direct primary handoffs, general={}",
            index == 1
        );
    }
}

type Capture = (usize, Rule, usize, usize, String);

fn capture_signature_inner<'a>(
    pairs: impl IntoIterator<Item = grameo::compat::Pair<'a, Rule>>,
) -> Vec<Capture> {
    fn visit(pair: grameo::compat::Pair<'_, Rule>, depth: usize, result: &mut Vec<Capture>) {
        result.push((
            depth,
            pair.as_rule(),
            pair.as_span().start(),
            pair.as_span().end(),
            pair.as_str().into(),
        ));
        for child in pair.into_inner() {
            visit(child, depth + 1, result);
        }
    }
    let mut result = Vec::new();
    for pair in pairs {
        visit(pair, 0, &mut result);
    }
    result
}

#[test]
fn composed_unit_capture_trees_and_rejections_match_the_reference_engine() {
    for source in PROGRAMS.iter().copied().chain([
        "/* \u{3bb} */ ; typedef int T; T f(T x){return x+1;} ; /* tail */",
        "int valid(void){return 1+2;} int broken(void){return 3+;}",
        "int valid(void){return 1;} static int broken[]={1,",
    ]) {
        let run = |general| {
            with_unit_engine(general, || {
                with_backend(Backend::Native, || {
                    profile_construction(|| {
                        parse_top_level_units(source, Default::default())
                            .map(capture_signature_inner)
                            .map_err(|error| format!("{error:?}"))
                    })
                })
            })
        };
        let (generated, generated_counts) = run(false);
        let (reference, reference_counts) = run(true);
        assert_eq!(generated, reference, "{source}");
        assert_engine_counts(&generated_counts, false);
        assert_engine_counts(&reference_counts, true);
        if generated.is_err() {
            assert!(
                generated_counts.source_pratt_recaptures > 0,
                "{generated_counts:?}"
            );
            assert!(
                reference_counts.source_pratt_recaptures > 0,
                "{reference_counts:?}"
            );
            let peg = with_backend(Backend::Peg, || {
                parse_captures(Rule::program, source, Default::default())
                    .map(|_| ())
                    .map_err(|error| format!("{error:?}"))
            });
            assert_eq!(
                generated.map(|_| ()),
                peg,
                "whole-program diagnostic fallback: {source}"
            );
        }
    }
}

#[test]
fn finite_and_custom_unit_budgets_keep_the_general_engine_with_both_builders() {
    for source in [
        "/* \u{3bb} */ int f(int x){int a[]={1,2};return a[x]+x*2;}",
        "/* \u{3bb} */ int f(int x){return x+;}",
    ] {
        let offset = "/* \u{3bb} */ ".len();
        for max_steps in [0, 1, 50, 512, usize::MAX] {
            for max_rule_depth in [0, 8, grameo::MatchOptions::default().max_rule_depth] {
                let options = grameo::MatchOptions {
                    max_steps,
                    max_rule_depth,
                };
                let run = |general| {
                    with_unit_engine(general, || {
                        with_backend(Backend::Native, || {
                            let mut builder = owned::SourcePrattBuilder::native(source);
                            profile_construction(|| {
                                direct::build_unit_captures(source, true, |direct| {
                                    build_hybrid_unit_inner(
                                        Rule::frontend_unit,
                                        source,
                                        offset,
                                        options,
                                        &mut WithDirect::new(&mut builder, direct),
                                    )
                                })
                                .map(|tree| {
                                    (
                                        tree.consumed(),
                                        capture_signature_inner(tree.into_typed_pairs::<Rule>()),
                                    )
                                })
                                .map_err(|error| format!("{error:?}"))
                            })
                        })
                    })
                };
                let (actual, counts) = run(false);
                let (expected, reference_counts) = run(true);
                assert_eq!(actual, expected, "{options:?}: {source}");
                let custom = max_steps != usize::MAX
                    || max_rule_depth != grameo::MatchOptions::default().max_rule_depth;
                assert_engine_counts(&counts, custom);
                assert_engine_counts(&reference_counts, true);
                assert_eq!(
                    counts.general_unit_matches + counts.generated_unit_matches,
                    1
                );
                if let Ok((end, captures)) = &actual {
                    assert_eq!(*end, source.len());
                    assert_eq!(captures[0].2, offset);
                    assert!(
                        captures
                            .iter()
                            .all(|(_, _, start, end, _)| *start >= offset && *end <= source.len())
                    );
                }
            }
        }
    }
}

#[test]
fn custom_parser_budgets_preserve_the_whole_program_policy() {
    let source = "int f(int x){int a[]={1,2};return a[x]+x*2;} int main(void){return f(0);}";
    for options in [
        grameo::MatchOptions {
            max_steps: 0,
            ..Default::default()
        },
        grameo::MatchOptions {
            max_steps: 100,
            ..Default::default()
        },
        grameo::MatchOptions {
            max_rule_depth: 8,
            ..Default::default()
        },
    ] {
        let run = |general| {
            with_unit_engine(general, || {
                with_backend(Backend::Native, || {
                    profile_matching(|| {
                        profile_construction(|| {
                            parse_top_level_units(source, options)
                                .map(capture_signature_inner)
                                .map_err(|error| format!("{error:?}"))
                        })
                    })
                })
            })
        };
        let ((actual, counts), matching) = run(false);
        let ((reference, reference_counts), _) = run(true);
        assert_eq!(actual, reference, "{options:?}");
        for counts in [counts, reference_counts] {
            assert_eq!(counts.generated_unit_matches, 0, "{counts:?}");
            assert_eq!(counts.general_unit_matches, 0, "{counts:?}");
            assert_eq!(counts.source_pratt_matches, 0, "{counts:?}");
        }
        assert!(
            matching
                .rules
                .iter()
                .any(|row| row.rule == Rule::program.name()),
            "{matching:?}"
        );
        assert!(
            !matching
                .rules
                .iter()
                .any(|row| row.rule == Rule::frontend_unit.name()),
            "{matching:?}"
        );
    }
}

#[test]
fn indexed_callback_comparisons_retain_the_declared_numeric_result_type() {
    crate::register();
    for (name, parse) in [
        (
            "peg",
            crate::parse_with_peg as vybe_runtime::registry::ParseFn,
        ),
        (
            "generated",
            parse_generated as vybe_runtime::registry::ParseFn,
        ),
        ("general", parse_general as vybe_runtime::registry::ParseFn),
    ] {
        let module = parse(PROGRAMS[4]).unwrap();
        let main = module
            .body
            .iter()
            .find(|statement| {
                matches!(&statement.kind,
            vybe_ast::StmtKind::FunctionDecl { name, .. } if name == "main")
            })
            .unwrap();
        let profile = vybe_compiler::profile::parse_profile(crate::profile_source()).unwrap();
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile_with_imports(&module)
            .unwrap()
            .chunks;
        let chunk = chunks.iter().find(|chunk| chunk.name == "main").unwrap();
        let dump = vybe_runtime::debug::disassemble(chunk);
        assert!(!dump.contains("dyneq"), "parser={name}\n{dump}\n{main:#?}");
    }
}

#[test]
fn generated_and_general_units_preserve_standard_wasm_owned_emission_vm_and_noeq() {
    crate::register();
    // Windowed FunctionExpr normalization is covered above, not launched in a
    // headless compiler test. All other fixtures must execute with exit code 0.
    for (engine, parse) in [
        (
            "generated",
            parse_generated as vybe_runtime::registry::ParseFn,
        ),
        ("general", parse_general as vybe_runtime::registry::ParseFn),
    ] {
        for (index, source) in PROGRAMS.iter().copied().enumerate() {
            eprintln!("unit_engine standard_wasm engine={engine} case={index}: {source}");
            unit_compatibility::execute_standard_wasm_programs(
                crate::profile_source(),
                [crate::parse_with_peg, parse],
                [source],
            );
        }
    }
}
