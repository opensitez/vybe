use super::*;
use vybe_ast::*;
use super::super::{merge_partial_classes, merge_partial_classes_whole_module_oracle};

fn stmt(kind: StmtKind) -> Statement { Statement::new(kind) }
fn marker(value: i64) -> Statement { stmt(StmtKind::Expr(Expression::int(value))) }
fn field(name: &str, value: i64) -> ClassMember {
    ClassMember::Field { name: name.into(), type_hint: Some("int32".into()),
        init: Some(Expression::int(value)), modifiers: Default::default(), with_events: false,
        array_bounds: None, storage: None }
}
fn constructor(value: i64) -> ClassMember {
    ClassMember::Constructor { name: None, params: Vec::new(), body: vec![marker(value)],
        base_args: Some(vec![Expression::int(value)]), initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public }
}
fn class(name: &str, members: Vec<ClassMember>, partial: bool, parents: &[&str], interfaces: &[&str]) -> Statement {
    stmt(StmtKind::ClassDecl { name: name.into(), parents: parents.iter().map(|name| (*name).into()).collect(),
        interfaces: interfaces.iter().map(|name| (*name).into()).collect(), members,
        modifiers: ClassModifiers { is_partial: partial, ..Default::default() },
        decorators: vec![Expression::int(if partial { 1 } else { 2 })] })
}

fn assert_splits(body: &[Statement]) {
    let original = format!("{body:?}");
    for case_sensitive in [false, true] {
        let expected = format!("{:?}", merge_partial_classes_whole_module_oracle(body, case_sensitive));
        assert_eq!(format!("{:?}", merge_partial_classes(body, case_sensitive)), expected);
        for mask in 0..(1usize << body.len().saturating_sub(1)) {
            let mut preparation = PartialClassPreparation::new(case_sensitive, 0);
            preparation.consume(&[]);
            let mut start = 0;
            for end in 1..body.len() {
                if mask & (1 << (end - 1)) != 0 {
                    preparation.consume(&body[start..end]);
                    start = end;
                }
            }
            preparation.consume(&body[start..]);
            preparation.consume(&[]);
            assert_eq!(format!("{:?}", preparation.finish()), expected,
                "partition={mask} case_sensitive={case_sensitive}");
        }
    }
    assert_eq!(format!("{body:?}"), original);
}

#[test]
fn staged_declarations_preserve_first_position_metadata_spans_and_forward_references() {
    let mut body = vec![
        stmt(StmtKind::Expr(Expression::new(ExprKind::New {
            class: Box::new(Expression::ident("Future")), args: Vec::new() }))),
        class("Grouped", vec![field("first", 1)], true, &["Base"], &["First"]),
        marker(3),
        class("Future", vec![field("future", 4)], false, &["Grouped"], &[]),
        class("Grouped", vec![constructor(5), field("last", 6)], false, &["Other"], &["Last"]),
        marker(7),
    ];
    for (index, statement) in body.iter_mut().enumerate() {
        statement.span = Span { start_line: index as u32 + 1, start_col: 2,
            end_line: index as u32 + 1, end_col: 17 };
    }
    let merged = merge_partial_classes(&body, true);
    assert_eq!(merged.len(), 5);
    assert_eq!(merged[1].span.start_line, 2);
    let StmtKind::ClassDecl { members, parents, interfaces, modifiers, decorators, .. } = &merged[1].kind
        else { panic!("first declaration moved") };
    assert_eq!(members.len(), 3);
    assert_eq!(parents, &["Base", "Other"]);
    assert_eq!(interfaces, &["First", "Last"]);
    assert!(modifiers.is_partial);
    assert_eq!(format!("{decorators:?}"), format!("{:?}", vec![Expression::int(1)]));
    assert_splits(&body);
}

#[test]
fn staged_declarations_preserve_constructor_suppression_and_initial_duplicates() {
    for first in [vec![field("first", 1)], vec![constructor(1), constructor(2)]] {
        let body = vec![class("Same", first.clone(), true, &[], &[]),
            class("Same", vec![constructor(3), constructor(4), field("second", 2)], true, &[], &[]),
            class("Same", vec![constructor(5), field("third", 3)], false, &[], &[])];
        let merged = merge_partial_classes(&body, true);
        let StmtKind::ClassDecl { members, .. } = &merged[0].kind else { unreachable!() };
        let original = first.iter().filter(|member| matches!(member, ClassMember::Constructor { .. })).count();
        assert_eq!(members.iter().filter(|member| matches!(member, ClassMember::Constructor { .. })).count(), original.max(1));
        assert_splits(&body);
    }
}

#[test]
fn staged_declarations_preserve_case_folding_and_parent_interface_order() {
    let body = vec![
        class("Widget", vec![], true, &["Base", "Base", "\u{c4}Parent"], &["IFirst"]),
        class("WIDGET", vec![field("later", 2)], false, &["base", "\u{e4}parent", "Second"], &["ifirst", "ISecond"]),
        class("\u{c4}Type", vec![field("upper", 3)], true, &[], &[]),
        class("\u{e4}type", vec![field("lower", 4)], false, &[], &[]),
    ];
    let merged = merge_partial_classes(&body, false);
    assert_eq!(merged.len(), 2);
    let StmtKind::ClassDecl { parents, interfaces, .. } = &merged[0].kind else { unreachable!() };
    assert_eq!(parents, &["Base", "Base", "\u{c4}Parent", "Second"]);
    assert_eq!(interfaces, &["IFirst", "ISecond"]);
    assert_eq!(merge_partial_classes(&body, true).len(), 4);
    assert_splits(&body);
}

#[test]
fn staged_declarations_do_not_merge_nested_namespaces_or_different_declaration_kinds() {
    let nested = vec![class("Same", vec![field("nested", 1)], true, &[], &[]),
        class("Same", vec![field("also_nested", 2)], true, &[], &[])];
    let body = vec![class("Same", vec![], true, &[], &[]),
        stmt(StmtKind::NamespaceDecl { name: "scope".into(), body: nested.clone() }),
        stmt(StmtKind::Block(nested)),
        stmt(StmtKind::StructDecl { name: "Same".into(), interfaces: Vec::new(),
            members: vec![field("struct_field", 4)], visibility: Visibility::Public,
            decorators: Vec::new(), semantics: Default::default() }),
        marker(4), class("Same", vec![field("top", 5)], false, &[], &[])];
    let merged = merge_partial_classes(&body, true);
    assert_eq!(merged.len(), 5);
    assert_eq!(format!("{:?}", merged[1]), format!("{:?}", body[1]));
    assert_eq!(format!("{:?}", merged[2]), format!("{:?}", body[2]));
    assert_eq!(format!("{:?}", merged[3]), format!("{:?}", body[3]));
    assert_splits(&body);
}

#[test]
fn staged_declaration_output_owns_units_and_abandoned_merges_are_isolated() {
    let body = vec![class("Same", vec![field("first", 1)], true, &[], &[]),
        marker(2), class("Same", vec![field("last", 3)], false, &[], &[])];
    let expected = format!("{:?}", merge_partial_classes_whole_module_oracle(&body, true));
    let mut preparation = PartialClassPreparation::new(true, 0);
    for statement in body { preparation.consume(std::slice::from_ref(&statement)); }
    assert_eq!(format!("{:?}", preparation.finish()), expected);
    let mut abandoned = PartialClassPreparation::new(true, 0);
    abandoned.consume(&[class("abandoned", vec![], true, &[], &[])]);
    drop(abandoned);
    assert!(PartialClassPreparation::new(true, 0).finish().is_empty());
    assert_splits(&[]);
}
