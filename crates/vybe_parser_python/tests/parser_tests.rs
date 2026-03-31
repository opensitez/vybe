use vybe_parser_python::{parse, Stmt, Expr};

#[test]
fn parse_while_if_elif_break_continue_normalization() {
    let src = r#"
while True:
  if i == 3:
    continue
  elif i == 4:
    break
"#;
    let prog = parse(src).expect("parse failed");
    // Expect one statement (while)
    assert_eq!(prog.stmts.len(), 1);
    match &prog.stmts[0] {
        Stmt::While { cond: _, body } => {
            // body should contain one If
            assert!(!body.is_empty());
            match &body[0] {
                Stmt::If { cond: _, then_branch, else_branch } => {
                    // then_branch first stmt should be Continue (normalized)
                    assert!(matches!(then_branch.get(0), Some(Stmt::Continue)));
                    // elif should have produced nested If in else_branch with Break
                    if let Some(eb) = else_branch {
                        // either nested If or vector containing it
                        let found_break = eb.iter().any(|s| matches!(s, Stmt::If { .. } ) );
                        assert!(found_break, "expected nested elif If node");
                    } else {
                        panic!("expected else_branch for elif");
                    }
                }
                other => panic!("expected If inside while body, got: {:?}", other),
            }
        }
        other => panic!("expected While as top stmt, got: {:?}", other),
    }
}

#[test]
fn test_list_parsing() {
    let src = r#"
lst = [1, 2, "three"]
"#;
    let prog = vybe_parser_python::parse(src).expect("parse failed");
    assert_eq!(prog.stmts.len(), 1);
    match &prog.stmts[0] {
        Stmt::Assign { name, expr } => {
            assert_eq!(name, "lst");
            if let Expr::List(items) = expr { assert_eq!(items.len(), 3); } else { panic!("expected list"); }
        }
        _ => panic!("expected assign for lst"),
    }
}
#[test]
fn test_normalize_break_continue_in_nested_ifs() {
    let src = r#"
 i = 0
 while i < 6:
   i = i + 1
   if i == 3:
     continue
   if i == 5:
     break
   print(i)
"#;
    let prog = parse(src).expect("parse failed");
    // find the while stmt
    let mut found = false;
    for s in prog.stmts.iter() {
        if let Stmt::While { cond: _, body } = s {
            // expect there to be an If whose then_branch first element is Continue
            for bs in body.iter() {
                if let Stmt::If { cond: _, then_branch, else_branch: _ } = bs {
                    if !then_branch.is_empty() {
                        if let Stmt::Continue = then_branch[0] {
                            found = true;
                        }
                        if let Stmt::Break = then_branch[0] {
                            found = true;
                        }
                    }
                }
            }
        }
    }
    assert!(found, "expected Break/Continue to be normalized into Stmt variants");
}

#[test]
fn test_logical_and_or_parsing() {
    let src = r#"
if 1 and 0:
  print("and")
if 1 or 0:
  print("or")
"#;
    let prog = parse(src).expect("parse failed");
    let mut saw_and_or = (false, false);
    for s in prog.stmts.iter() {
        if let Stmt::If { cond, then_branch: _, else_branch: _ } = s {
            match cond {
                Expr::Binary { op, .. } if op == "and" => saw_and_or.0 = true,
                Expr::Binary { op, .. } if op == "or" => saw_and_or.1 = true,
                _ => {}
            }
        }
    }
    assert!(saw_and_or.0 && saw_and_or.1, "expected to see both and/or binary exprs");
}

#[test]
fn test_tuple_and_dict_parsing() {
    let src = r#"
x = (1, 2, 3)
y = {"a": 1, b: 2}
"#;
    let prog = vybe_parser_python::parse(src).expect("parse failed");
    // Expect two assigns
    assert_eq!(prog.stmts.len(), 2);
    match &prog.stmts[0] {
        Stmt::Assign { name, expr } => {
            assert_eq!(name, "x");
            if let Expr::Tuple(e) = expr { assert_eq!(e.len(), 3); } else { panic!("expected tuple"); }
        }
        _ => panic!("expected assign for x"),
    }
    match &prog.stmts[1] {
        Stmt::Assign { name, expr } => {
            assert_eq!(name, "y");
            if let Expr::Dict(items) = expr { assert_eq!(items.len(), 2); } else { panic!("expected dict"); }
        }
        _ => panic!("expected assign for y"),
    }
}
