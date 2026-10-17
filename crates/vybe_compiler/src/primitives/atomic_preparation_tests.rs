use super::*;
use vybe_ast::*;

fn expr(kind: ExprKind) -> Expression { Expression::new(kind) }
fn ident(name: &str) -> Box<Expression> { Box::new(expr(ExprKind::Ident(name.into()))) }
fn stmt(kind: StmtKind) -> Statement { Statement::new(kind) }
fn atomic(op: AtomicOp) -> Statement { stmt(StmtKind::Expr(expr(ExprKind::Atomic(op)))) }
fn load(name: &str) -> Statement {
    atomic(AtomicOp::Load { place: ident(name), ordering: MemoryOrder::SeqCst })
}
fn function(name: &str, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::FunctionDecl { name: name.into(), params: Vec::new(), return_type: None,
        body, modifiers: Default::default(), handles: Vec::new(),
        is_async: false, is_generator: false, is_sub: false })
}
fn class(name: &str, members: Vec<ClassMember>) -> Statement {
    stmt(StmtKind::ClassDecl { name: name.into(), parents: Vec::new(), interfaces: Vec::new(),
        members, modifiers: Default::default(), decorators: Vec::new() })
}

fn oracle(body: &[Statement], mut names: HashSet<String>, mut receipts: HashSet<(usize, usize)>)
    -> (HashSet<String>, HashSet<(usize, usize)>) {
    // Original whole-module call and original negative-receipt invalidation.
    collect_atomic_place_idents(body, &mut AtomicModuleScan {
        class_depth: 0, names: &mut names, bodies: &mut receipts,
    });
    if !names.is_empty() { receipts.clear(); }
    (names, receipts)
}

fn assert_splits(body: &[Statement], seed: HashSet<String>) {
    let original = format!("{body:?}");
    let initial_receipts = HashSet::from([(17, 2)]);
    let expected = oracle(body, seed.clone(), initial_receipts.clone());
    for mask in 0..(1usize << body.len().saturating_sub(1)) {
        for keep_source_receipts in [false, true] {
            let mut preparation = ModuleAtomicPreparation::new(seed.clone());
            let mut receipts = initial_receipts.clone();
            let mut consume = |statements: &[Statement]| preparation.consume(statements,
                keep_source_receipts.then_some(&mut receipts));
            consume(&[]);
            let mut start = 0;
            for end in 1..body.len() {
                if mask & (1 << (end - 1)) != 0 {
                    consume(&body[start..end]);
                    start = end;
                }
            }
            consume(&body[start..]);
            consume(&[]);
            let names = preparation.finish();
            assert_eq!(names, expected.0, "partition={mask}");
            if keep_source_receipts {
                if !names.is_empty() { receipts.clear(); }
                assert_eq!(receipts, expected.1, "partition={mask}");
            } else {
                assert_eq!(receipts, initial_receipts, "owned-facts mode retained source addresses");
            }
        }
    }
    assert_eq!(format!("{body:?}"), original);
}

#[test]
fn staged_atomic_preparation_preserves_negative_receipts_and_class_boundaries() {
    let body = vec![
        function("outer", vec![function("nested", vec![stmt(StmtKind::Empty)])]),
        class("copied", vec![ClassMember::Method(Box::new(function("method", vec![
            function("nested_method", vec![]),
        ])))]),
        function("after_class", vec![stmt(StmtKind::Empty)]),
        atomic(AtomicOp::Fence { ordering: MemoryOrder::Acquire }),
    ];
    let (names, receipts) = oracle(&body, HashSet::new(), HashSet::new());
    assert!(names.is_empty());
    assert_eq!(receipts.len(), 3, "class procedures must never receive source-storage proofs");
    assert_splits(&body, HashSet::new());
    assert_splits(&body, HashSet::from(["previous_unit_atomic".into()]));
}

#[test]
fn staged_atomic_preparation_preserves_all_atomic_operands_and_late_invalidation() {
    let body = vec![
        function("early", vec![stmt(StmtKind::Empty)]),
        atomic(AtomicOp::Store { place: ident("store"), value: ident("value_not_place"), ordering: MemoryOrder::Release }),
        atomic(AtomicOp::Rmw { op: AtomicRmw::Add, place: ident("rmw"), operand: ident("operand_not_place"),
            result: RmwResult::New, ordering: MemoryOrder::AcqRel }),
        atomic(AtomicOp::CompareExchange { place: ident("compare"), expected: ident("expected_not_place"),
            replacement: Box::new(expr(ExprKind::Atomic(AtomicOp::Load {
                place: ident("nested_load"), ordering: MemoryOrder::Acquire,
            }))), result: RmwResult::Old, ordering: MemoryOrder::SeqCst }),
        function("late", vec![load("load")]),
    ];
    let (names, receipts) = oracle(&body, HashSet::new(), HashSet::new());
    assert_eq!(names, ["store", "rmw", "compare", "nested_load", "load"].map(str::to_owned).into());
    assert!(receipts.is_empty(), "a late atomic place must revoke earlier negative receipts");
    assert_splits(&body, HashSet::new());
}

#[test]
fn staged_atomic_preparation_preserves_nested_callable_and_class_places() {
    let body = vec![
        stmt(StmtKind::Expr(expr(ExprKind::Lambda { params: Vec::new(),
            body: LambdaBody::Block(vec![load("lambda_place")]), is_async: false, captures: Vec::new() }))),
        stmt(StmtKind::Expr(expr(ExprKind::FunctionExpr(Box::new(function("value", vec![load("function_place")])))))),
        class("container", vec![ClassMember::Method(Box::new(function("method", vec![load("class_place")])))]),
        stmt(StmtKind::NamespaceDecl { name: "namespace".into(), body: vec![
            stmt(StmtKind::Block(vec![load("namespace_place")])),
        ] }),
    ];
    assert_eq!(oracle(&body, HashSet::new(), HashSet::new()).0,
        ["lambda_place", "function_place", "class_place", "namespace_place"].map(str::to_owned).into());
    assert_splits(&body, HashSet::new());
}

#[test]
fn staged_atomic_preparation_keeps_complex_places_out_of_root_name_facts() {
    let body = vec![atomic(AtomicOp::Load { place: Box::new(expr(ExprKind::Index {
        object: ident("container_not_binding"), index: Box::new(expr(ExprKind::Atomic(AtomicOp::Load {
            place: ident("index_place"), ordering: MemoryOrder::Relaxed,
        }))), null_safe: false,
    })), ordering: MemoryOrder::SeqCst })];
    assert_eq!(oracle(&body, HashSet::new(), HashSet::new()).0, HashSet::from(["index_place".into()]));
    assert_splits(&body, HashSet::new());
}

#[test]
fn staged_atomic_facts_are_owned_and_abandoned_sources_do_not_publish() {
    fn owned_and_send<T: Send + 'static>() {}
    owned_and_send::<ModuleAtomicPreparation>();
    let body = vec![load("owned_place"), function("f", vec![load("other_place")])];
    let mut preparation = ModuleAtomicPreparation::new(HashSet::new());
    for statement in body { preparation.consume(std::slice::from_ref(&statement), None); }
    assert_eq!(preparation.finish(), ["owned_place", "other_place"].map(str::to_owned).into());
    let mut abandoned = ModuleAtomicPreparation::new(HashSet::new());
    abandoned.consume(&[load("failed_source")], None);
    drop(abandoned);
    assert!(ModuleAtomicPreparation::new(HashSet::new()).finish().is_empty());
    assert_splits(&[], HashSet::new());
}
