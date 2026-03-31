use vybe_parser_python::parse;
use vybe_compiler_python::Compiler;

#[test]
fn test_compile_loop_demo() {
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
    let c = Compiler::new();
    let res = c.compile(&prog);
    assert!(res.is_ok(), "compile failed: {:?}", res.err());
    let chunks = res.unwrap();
    assert!(!chunks.is_empty(), "expected at least one chunk");
}

#[test]
fn test_compile_logical_ops() {
    let src = r#"
if 1 and 0:
  print("and")
if 1 or 0:
  print("or")
"#;
    let prog = parse(src).expect("parse failed");
    let c = Compiler::new();
    let res = c.compile(&prog);
    assert!(res.is_ok(), "compile failed: {:?}", res.err());
}

#[test]
fn test_compile_tuple_and_dict() {
  let src = r#"
tp = (1, 2)
dm = {"one": 1, "two": 2}
"#;
  let prog = parse(src).expect("parse failed");
  let c = Compiler::new();
  let res = c.compile(&prog);
  assert!(res.is_ok(), "compile failed: {:?}", res.err());
}

#[test]
fn test_compile_list_literal() {
  let src = r#"
lst = [1, 2, "three"]
"#;
  let prog = parse(src).expect("parse failed");
  let c = Compiler::new();
  let res = c.compile(&prog);
  assert!(res.is_ok(), "compile failed: {:?}", res.err());
}
