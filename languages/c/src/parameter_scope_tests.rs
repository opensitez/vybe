use super::*;

fn parameter(name: &str) -> Param {
    Param {
        name: name.into(), type_hint: Some("int".into()), default: None,
        pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
        is_optional: false, is_nullable: false,
    }
}

#[test]
fn visible_parameter_names_match_full_snapshot_without_copying_unrelated_names() {
    for count in [0, 1, 17, 4096] {
        let mut types = (0..count).map(|index|
            (format!("visible_{index}"), "int".to_string())).collect::<HashMap<_, _>>();
        types.insert("shared".into(), "int".into());
        let old_snapshot = types.keys().cloned().collect::<HashSet<_>>();
        for names in [vec![], vec!["argument"], vec!["shared"],
            vec!["Shared", "shared", "argument", "shared"]] {
            let params = names.iter().map(|name| parameter(name)).collect::<Vec<_>>();
            let actual = visible_parameter_names(&params, &types);
            let expected = params.iter().filter(|param| old_snapshot.contains(&param.name))
                .map(|param| param.name.clone()).collect::<HashSet<_>>();
            assert_eq!(actual, expected);
            assert!(actual.len() <= params.len());
            let old_order = params.iter().filter(|param| old_snapshot.contains(&param.name))
                .map(|param| &param.name).collect::<Vec<_>>();
            let new_order = params.iter().filter(|param| actual.contains(&param.name))
                .map(|param| &param.name).collect::<Vec<_>>();
            assert_eq!(new_order, old_order);
        }
    }
}

const SOURCE: &str = "static int shared=7;int f(int shared){return shared+1;} \
    int g(void){return shared;}int main(void){return f(2)!=3 || g()!=7;}";

#[test]
fn parameter_collision_renames_only_the_parameter_binding() {
    for parse in [crate::parse_with_peg, crate::parse_optimized] {
        let module = parse(SOURCE).unwrap();
        let mut parameter_name = None;
        let mut global_read = false;
        for statement in &module.body {
            if let StmtKind::FunctionDecl { name, params, body, .. } = &statement.kind {
                if name == "f" {
                    assert_ne!(params[0].name, "shared");
                    let renamed = &params[0].name;
                    let mut reads_parameter = false;
                    for statement in body {
                        statement.walk_exprs(&mut |value| {
                            if let ExprKind::Ident(name) = &value.kind {
                                assert_ne!(name, "shared");
                                reads_parameter |= name == renamed;
                            }
                        });
                    }
                    assert!(reads_parameter);
                    parameter_name = Some(renamed.clone());
                } else if name == "g" {
                    for statement in body {
                        statement.walk_exprs(&mut |value|
                            global_read |= matches!(&value.kind, ExprKind::Ident(name) if name == "shared"));
                    }
                }
            }
        }
        assert!(parameter_name.is_some() && global_read);
    }
}

#[allow(dead_code)]
#[path = "../tests/support/frontend_unit_compatibility.rs"]
mod unit_compatibility;

#[test]
fn parameter_collision_executes_standard_wasm_without_dynamic_equality() {
    crate::register();
    unit_compatibility::execute_standard_wasm_programs(crate::profile_source(),
        [crate::parse_with_peg, crate::parse_optimized], [SOURCE]);
}
