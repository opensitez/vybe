use vybec::parser_python::ast::*;

fn parse(src: &str) -> Module {
    vybec::parser_python::parse(src).expect("parse failed")
}

fn first_value(m: &Module) -> &Expression {
    match &m.body[0] {
        Statement::Assign { value, .. } => value,
        other => panic!("expected Assign, got: {:?}", other),
    }
}

#[test]
fn list_comp_basic() {
    let m = parse("x = [i for i in items]\n");
    match first_value(&m) {
        Expression::ListComp { generators, .. } => {
            assert_eq!(generators.len(), 1);
            assert!(generators[0].ifs.is_empty());
        }
        other => panic!("expected ListComp, got: {:?}", other),
    }
}

#[test]
fn list_comp_with_filter() {
    let m = parse("x = [i for i in items if i > 0]\n");
    match first_value(&m) {
        Expression::ListComp { generators, .. } => {
            assert_eq!(generators[0].ifs.len(), 1);
        }
        other => panic!("expected ListComp, got: {:?}", other),
    }
}

#[test]
fn list_comp_with_expression() {
    let m = parse("x = [i * 2 for i in range(10)]\n");
    match first_value(&m) {
        Expression::ListComp { element, .. } => {
            assert!(matches!(element.as_ref(), Expression::BinOp { op: BinOp::Mul, .. }));
        }
        other => panic!("expected ListComp, got: {:?}", other),
    }
}

#[test]
fn list_comp_nested_for() {
    let m = parse("x = [i + j for i in a for j in b]\n");
    match first_value(&m) {
        Expression::ListComp { generators, .. } => {
            assert_eq!(generators.len(), 2);
        }
        other => panic!("expected ListComp, got: {:?}", other),
    }
}

#[test]
fn set_comp() {
    let m = parse("x = {i for i in items}\n");
    assert!(matches!(first_value(&m), Expression::SetComp { .. }));
}

#[test]
fn dict_comp() {
    let m = parse("x = {k: v for k, v in items}\n");
    assert!(matches!(first_value(&m), Expression::DictComp { .. }));
}

#[test]
fn generator_expr_in_call() {
    let m = parse("x = sum(i for i in range(10))\n");
    match first_value(&m) {
        Expression::Call { args, .. } => {
            assert_eq!(args.len(), 1);
            assert!(matches!(&args[0], Expression::GeneratorExp { .. }));
        }
        other => panic!("expected Call, got: {:?}", other),
    }
}
