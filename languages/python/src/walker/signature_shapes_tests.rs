use super::*;

fn define(w: &mut PyWalker, name: &str, params: Vec<Param>, kinds: Option<Vec<String>>) {
    w.py_defined_functions.insert(name.into());
    w.py_defined_function_params.insert(name.into(), params);
    if let Some(kinds) = kinds { w.py_defined_function_param_kinds.insert(name.into(), kinds); }
}
fn compare(w: &mut PyWalker, name: &str, args: Vec<Argument>) {
    let callee = Expression::ident(name);
    let expected = rewrite_python_known_callable_call(w, &callee, args.clone(), false);
    let actual = rewrite(w, &callee, &args, false);
    assert_eq!(format!("{expected:?}"), format!("{actual:?}"));
}
#[test]
fn signature_metadata_stays_live_across_alias_rebinding_and_definition_priority() {
    let mut w = PyWalker::default();
    define(&mut w, "f", vec![plain_param("a", Some(Expression::int(1)))], None);
    w.py_function_aliases.insert("g".into(), "f".into());
    compare(&mut w, "g", vec![]);
    w.py_defined_function_params.insert("f".into(), vec![plain_param("a", None)]);
    compare(&mut w, "g", vec![]);
    define(&mut w, "g", vec![plain_param("b", Some(Expression::int(2)))], None);
    compare(&mut w, "g", vec![]);
    w.py_defined_functions.remove("g");
    compare(&mut w, "g", vec![]);
    w.py_function_aliases.insert("g".into(), "h".into());
    w.py_function_aliases.insert("h".into(), "g".into());
    compare(&mut w, "g", vec![]);
}
#[test]
fn missing_empty_and_default_kind_metadata_keep_original_results() {
    let mut w = PyWalker::default();
    w.py_defined_functions.insert("f".into());
    w.py_defined_function_param_kinds.insert("f".into(), vec!["KEYWORD_ONLY".into()]);
    compare(&mut w, "f", vec![Argument::positional(Expression::int(1))]);
    w.py_defined_function_params.insert("f".into(), vec![]);
    compare(&mut w, "f", vec![Argument::positional(Expression::int(1))]);
    w.py_defined_function_params.insert("f".into(), vec![plain_param("a", None)]);
    w.py_defined_function_param_kinds.remove("f");
    compare(&mut w, "f", vec![Argument::positional(Expression::int(1))]);
    w.py_defined_function_param_kinds.insert("f".into(), vec!["KEYWORD_ONLY".into()]);
    compare(&mut w, "f", vec![Argument::positional(Expression::int(1))]);
}
#[test]
fn descriptor_error_order_and_packing_fallback_match_owned_reference() {
    let mut w = PyWalker::default();
    define(&mut w, "f", vec![plain_param("a", None), plain_param("b", None)], None);
    let mut named = Argument::positional(Expression::int(2)); named.name = Some("a".into());
    compare(&mut w, "f", vec![Argument::positional(Expression::int(1)), named]);
    compare(&mut w, "f", vec![]);
    let mut spread = Argument::positional(Expression::new(ExprKind::Tuple(vec![Expression::int(1), Expression::int(2)]))); spread.spread = true;
    compare(&mut w, "f", vec![spread]);
    w.py_defined_function_param_kinds.insert("f".into(), vec!["VAR_POSITIONAL".into(), "KEYWORD_ONLY".into()]);
    compare(&mut w, "f", vec![Argument::positional(Expression::int(1)), Argument::positional(Expression::int(2))]);
}
#[test]
fn effectful_looking_defaults_and_arguments_are_not_evaluated_for_discarded_validation() {
    let mut w = PyWalker::default();
    define(&mut w, "f", vec![plain_param("a", Some(call_ident("unknown_default", vec![])))], None);
    compare(&mut w, "f", vec![]);
    compare(&mut w, "f", vec![Argument::positional(call_ident("unknown_argument", vec![]))]);
}
