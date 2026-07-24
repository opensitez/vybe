use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Category 3: Dynamic Interactive Code Compilation (codeop module)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_codeop_compile_complete_statement() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("x = 42\nprint(x)")
print(code is not None)
exec(code)
"#);
    assert_eq!(out, vec!["True", "42"]);
}

#[test]
fn test_codeop_compile_incomplete_statement() {
    let out = run_python(r#"
import codeop
# Incomplete def statement returns None
code = codeop.compile_command("def foo():")
print(code is None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_compile_syntax_error() {
    let out = run_python(r#"
import codeop
try:
    codeop.compile_command("def foo() pass")
except SyntaxError:
    print("SyntaxErrorCaught")
"#);
    assert_eq!(out, vec!["SyntaxErrorCaught"]);
}

#[test]
fn test_codeop_compile_command_eval_symbol() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("1 + 2", symbol="eval")
print(code is not None)
print(eval(code))
"#);
    assert_eq!(out, vec!["True", "3"]);
}

#[test]
fn test_codeop_compile_command_exec_symbol() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("a = 10; b = 20", symbol="exec")
env = {}
exec(code, env)
print(env['a'] + env['b'])
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_codeop_command_compiler_class() {
    let out = run_python(r#"
import codeop
compiler = codeop.CommandCompiler()
code = compiler("print('Compiled')")
print(code is not None)
exec(code)
"#);
    assert_eq!(out, vec!["True", "Compiled"]);
}

#[test]
fn test_codeop_compile_class() {
    let out = run_python(r#"
import codeop
compiler = codeop.Compile()
code = compiler("val = 100", symbol="exec")
env = {}
exec(code, env)
print(env['val'])
"#);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_codeop_incomplete_multiline_string() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command('s = """hello')
print(code is None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_incomplete_unclosed_parenthesis() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command('x = (1 + 2')
print(code is None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_empty_string() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("")
print(code is None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_comment_only() {
    let out = run_python(r##"
import codeop
code = codeop.compile_command("# Just a comment")
print(code is None)
"##);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_custom_filename() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("x = 1", filename="<custom_input>")
print(code.co_filename)
"#);
    assert_eq!(out, vec!["<custom_input>"]);
}

#[test]
fn test_codeop_incomplete_if_statement() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("if True:\n    pass\nelse:")
print(code is None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_complete_class_def() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("class MyClass:\n    pass\n")
print(code is not None)
env = {}
exec(code, env)
print('MyClass' in env)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_codeop_incomplete_try_except() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("try:\n    pass")
print(code is None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_overflow_error_handling() {
    let out = run_python(r#"
import codeop
try:
    # Overflow in literal float if compiler checks it
    codeop.compile_command("1e1000")
except OverflowError:
    print("OverflowErrorCaught")
except SyntaxError:
    print("SyntaxErrorCaught")
else:
    print("OK")
"#);
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn test_codeop_command_compiler_state_preservation() {
    let out = run_python(r#"
import codeop
cc = codeop.CommandCompiler()
c1 = cc("x = 10\n")
c2 = cc("y = 20\n")
print(c1 is not None and c2 is not None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_future_flags_compilation() {
    let out = run_python(r#"
import codeop, __future__
flags = __future__.annotations.compiler_flag
code = codeop.compile_command("x: int = 5", flags=flags)
print(code is not None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_incomplete_dict_literal() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("d = {'a': 1, 'b':")
print(code is None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_codeop_compile_decorator_incomplete() {
    let out = run_python(r#"
import codeop
code = codeop.compile_command("@decorator")
print(code is None)
"#);
    assert_eq!(out, vec!["True"]);
}
