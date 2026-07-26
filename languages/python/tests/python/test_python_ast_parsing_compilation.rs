use super::helpers::run_python;

#[test]
fn test_python_ast_parse_and_dump() {
    let src = r#"
import ast
expr = ast.parse('a + b * 2', mode='eval')
print(isinstance(expr, ast.Expression))
print('BinOp' in ast.dump(expr))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_python_compile_eval_codeobj() {
    let src = r#"
code = compile('x = 5\ny = x * 3\n', '<test>', 'exec')
ns = {}
exec(code, ns)
print(ns['y'])
"#;
    assert_eq!(run_python(src), vec!["15"]);
}
