use super::*;
use std::collections::BTreeMap;

fn function(name: &str, parameter: &str, return_type: Option<&str>, generator: bool) -> Statement {
    Statement::new(StmtKind::FunctionDecl { name: name.into(), params: vec![Param {
        name: parameter.into(), type_hint: Some("declared_type".into()), default: None,
        pass_by: PassBy::Value, is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false,
    }], return_type: return_type.map(str::to_owned), body: Vec::new(), modifiers: Default::default(),
        handles: Vec::new(), is_async: false, is_generator: generator, is_sub: false })
}

fn snapshot(compiler: &Compiler) -> String {
    fn map<V: std::fmt::Debug, S: std::hash::BuildHasher>(values: &HashMap<String, V, S>) -> BTreeMap<String, String> {
        values.iter().map(|(key, value)| (key.clone(), format!("{value:?}"))).collect()
    }
    let mut globals = compiler.defined_globals.iter().cloned().collect::<Vec<_>>();
    let mut functions = compiler.defined_functions.iter().cloned().collect::<Vec<_>>();
    let mut generators = compiler.generator_functions.iter().cloned().collect::<Vec<_>>();
    globals.sort(); functions.sort(); generators.sort();
    format!("{:?}", (globals, functions, generators,
        map(&compiler.function_param_modes), map(&compiler.function_param_types),
        map(&compiler.function_min_arity), map(&compiler.function_signatures),
        map(&compiler.function_return_types)))
}

#[test]
fn staged_function_registration_preserves_existing_entries_and_all_registry_tables() {
    let body = vec![function("Mixed", "first", None, false),
        function("mixed", "second", Some("second_return"), true),
        function("Mixed", "third", Some("third_return"), false),
        Statement::new(StmtKind::Empty), function("Forward", "argument", Some("Mixed"), false)];
    for case_sensitive in [false, true] {
        for buffered in [false, true] {
            for prior in [false, true] {
                let make = || {
                    let mut profile = crate::profile::parse_profile(
                        "[info]\nname='function-registration-test'\n[compiler]\n").unwrap();
                    profile.case_sensitive = case_sensitive;
                    profile.buffered_iterator_methods = buffered;
                    let mut compiler = Compiler::with_profile(profile);
                    if prior {
                        compiler.predeclare_function_names_whole_module_oracle(&[
                            function("Mixed", "prior_module", Some("prior_return"), false)]);
                    }
                    compiler
                };
                let mut expected = make();
                expected.predeclare_function_names_whole_module_oracle(&body);
                for mask in 0..(1usize << (body.len() - 1)) {
                    let mut actual = make();
                    let mut sealed = make();
                    let mut headers = super::super::function_preparation::FunctionRegistrationPreparation::default();
                    actual.predeclare_function_names(&[]);
                    let mut start = 0;
                    for end in 1..body.len() {
                        if mask & (1 << (end - 1)) != 0 {
                            actual.predeclare_function_names(&body[start..end]);
                            headers.consume(&body[start..end]);
                            start = end;
                        }
                    }
                    actual.predeclare_function_names(&body[start..]);
                    headers.consume(&body[start..]);
                    sealed.register_function_declarations(headers.finish());
                    assert_eq!(snapshot(&actual), snapshot(&expected),
                        "partition={mask} case_sensitive={case_sensitive} buffered={buffered} prior={prior}");
                    assert_eq!(snapshot(&sealed), snapshot(&expected),
                        "sealed partition={mask} case_sensitive={case_sensitive} buffered={buffered} prior={prior}");
                }
            }
        }
    }
}
