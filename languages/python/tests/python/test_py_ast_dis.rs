use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: ast + dis — AST parsing, unparse, code objects, disassembler
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_ast_parse_and_walk() {
    let src = r#"
import ast

code = "x = 10 + 20"
tree = ast.parse(code)

nodes = [type(n).__name__ for n in ast.walk(tree)]
print("Module" in nodes)
print("Assign" in nodes)
print("BinOp" in nodes)
print("Add" in nodes)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_py_ast_unparse() {
    let src = r#"
import ast

code = "x = (a + b) * c"
tree = ast.parse(code)
reconstructed = ast.unparse(tree).strip()
print(reconstructed)
"#;
    assert_eq!(run_python(src), vec!["x = (a + b) * c"]);
}

#[test]
fn test_py_ast_literal_eval() {
    let src = r#"
import ast

safe_dict = ast.literal_eval("{'a': 1, 'b': [2, 3, 4]}")
print(safe_dict["a"])
print(safe_dict["b"])

try:
    ast.literal_eval("__import__('os').system('ls')")
except ValueError:
    print("ValueError: unsafe code rejected")
"#;
    assert_eq!(
        run_python(src),
        vec!["1", "[2, 3, 4]", "ValueError: unsafe code rejected"]
    );
}

#[test]
fn test_py_ast_node_visitor() {
    let src = r#"
import ast

class FunctionCollector(ast.NodeVisitor):
    def __init__(self):
        self.names = []

    def visit_FunctionDef(self, node):
        self.names.append(node.name)
        self.generic_visit(node)

code = """
def foo():
    pass
def bar():
    def nested():
        pass
"""
tree = ast.parse(code)
collector = FunctionCollector()
collector.visit(tree)
print(collector.names)
"#;
    assert_eq!(run_python(src), vec!["['foo', 'bar', 'nested']"]);
}

#[test]
fn test_py_ast_node_transformer() {
    let src = r#"
import ast

class ConstantDoubler(ast.NodeTransformer):
    def visit_Constant(self, node):
        if isinstance(node.value, int):
            return ast.Constant(value=node.value * 2)
        return node

code = "x = 10 + 20"
tree = ast.parse(code)
transformed = ConstantDoubler().visit(tree)
ast.fix_missing_locations(transformed)
compiled = compile(transformed, filename="<ast>", mode="exec")

scope = {}
exec(compiled, scope)
print(scope["x"])  # 20 + 40 = 60
"#;
    assert_eq!(run_python(src), vec!["60"]);
}

#[test]
fn test_py_dis_bytecode_disassembly() {
    let src = r#"
import dis, io

def add(a, b):
    return a + b

buf = io.StringIO()
dis.dis(add, file=buf)
output = buf.getvalue()
print("BINARY_ADD" in output or "BINARY_OP" in output)
print("RETURN_VALUE" in output)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_dis_get_instructions() {
    let src = r#"
import dis

def multiply(a, b):
    return a * b

instructions = list(dis.get_instructions(multiply))
opnames = [inst.opname for inst in instructions]
print("LOAD_FAST" in opnames)
print("RETURN_VALUE" in opnames)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_code_object_attributes() {
    let src = r#"
def fn(x, y=10, *args, **kwargs):
    z = x + y
    return z

code = fn.__code__
print(code.co_argcount)
print("x" in code.co_varnames)
print("y" in code.co_varnames)
print("z" in code.co_varnames)
"#;
    assert_eq!(run_python(src), vec!["2", "True", "True", "True"]);
}

#[test]
fn test_py_compile_exec_eval_single() {
    let src = r#"
code_eval = compile("3 + 4", "<string>", "eval")
print(eval(code_eval))

code_exec = compile("a = 42\nb = a * 2", "<string>", "exec")
scope = {}
exec(code_exec, scope)
print(scope["b"])
"#;
    assert_eq!(run_python(src), vec!["7", "84"]);
}

#[test]
fn test_py_ast_dump() {
    let src = r#"
import ast

tree = ast.parse("x = 1")
dumped = ast.dump(tree)
print("Assign" in dumped)
print("targets" in dumped)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}
