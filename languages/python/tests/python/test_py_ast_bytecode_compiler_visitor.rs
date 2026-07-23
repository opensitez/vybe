use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: AST, Bytecode Compiler & Visitor — ast.parse, unparse, walk, literal_eval, NodeVisitor, NodeTransformer, dis.dis
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_ast_parse_walk_node_types() {
    let src = r#"
import ast

code = "def foo(a, b): return a + b"
tree = ast.parse(code)
node_types = [type(n).__name__ for n in ast.walk(tree)]

print("FunctionDef" in node_types)
print("Return" in node_types)
print("BinOp" in node_types)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_ast_unparse_reconstruction() {
    let src = r#"
import ast

expr = "y = x * 2 + 5"
tree = ast.parse(expr)
reconstructed = ast.unparse(tree).strip()
print(reconstructed)
"#;
    assert_eq!(run_python(src), vec!["y = x * 2 + 5"]);
}

#[test]
fn test_py_ast_literal_eval_safe_deserialization() {
    let src = r#"
import ast

parsed_list = ast.literal_eval("[1, 2, {'key': 'val'}]")
print(parsed_list[2]["key"])

try:
    ast.literal_eval("__import__('sys').exit(0)")
except ValueError:
    print("ValueError: unsafe expression rejected")
"#;
    assert_eq!(
        run_python(src),
        vec!["val", "ValueError: unsafe expression rejected"]
    );
}

#[test]
fn test_py_ast_node_visitor_custom_collector() {
    let src = r#"
import ast

class VariableCollector(ast.NodeVisitor):
    def __init__(self):
        self.vars = []
    def visit_Name(self, node):
        self.vars.append(node.id)
        self.generic_visit(node)

code = "x = a + b * c"
tree = ast.parse(code)
collector = VariableCollector()
collector.visit(tree)
print(sorted(list(set(collector.vars))))
"#;
    assert_eq!(run_python(src), vec!["['a', 'b', 'c', 'x']"]);
}

#[test]
fn test_py_ast_node_transformer_mutator() {
    let src = r#"
import ast

class IntIncrementer(ast.NodeTransformer):
    def visit_Constant(self, node):
        if isinstance(node.value, int):
            return ast.Constant(value=node.value + 1)
        return node

code = "x = 10"
tree = ast.parse(code)
transformed = IntIncrementer().visit(tree)
ast.fix_missing_locations(transformed)

compiled = compile(transformed, "<ast>", "exec")
scope = {}
exec(compiled, scope)
print(scope["x"])
"#;
    assert_eq!(run_python(src), vec!["11"]);
}

#[test]
fn test_py_dis_get_instructions_opnames() {
    let src = r#"
import dis

def add(x, y):
    return x + y

insts = list(dis.get_instructions(add))
opnames = [inst.opname for inst in insts]
print("LOAD_FAST" in opnames)
print("RETURN_VALUE" in opnames)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_code_object_co_consts_co_varnames() {
    let src = r#"
def fn(a, b):
    c = a + 10
    return c

code = fn.__code__
print(code.co_varnames)
print(10 in code.co_consts)
"#;
    assert_eq!(run_python(src), vec!["('a', 'b', 'c')", "True"]);
}

#[test]
fn test_py_compile_mode_eval_vs_exec() {
    let src = r#"
code_eval = compile("2 ** 3", "<eval>", "eval")
print(eval(code_eval))

code_exec = compile("res = [i for i in range(3)]", "<exec>", "exec")
scope = {}
exec(code_exec, scope)
print(scope["res"])
"#;
    assert_eq!(run_python(src), vec!["8", "[0, 1, 2]"]);
}

#[test]
fn test_py_ast_dump_string_representation() {
    let src = r#"
import ast

tree = ast.parse("x = 5")
dump = ast.dump(tree)
print("Assign" in dump)
print("Constant" in dump or "Num" in dump)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_dis_disassemble_code_to_string() {
    let src = r#"
import dis, io

def sample(): pass

buf = io.StringIO()
dis.dis(sample, file=buf)
out = buf.getvalue()
print("RETURN_VALUE" in out or "RETURN_CONST" in out)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
