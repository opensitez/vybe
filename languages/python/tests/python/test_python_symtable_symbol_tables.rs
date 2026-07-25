use super::helpers::run_python;

// symtable — symtable, SymbolTable, Symbol, get_type, get_name, is_referenced, is_imported, is_parameter, is_global, is_local, get_symbols, get_children

#[test]
fn test_symtable_parse_top_level_module() {
    let out = run_python(r#"
import symtable
st = symtable.symtable("x = 10; y = x + 5", "<string>", "exec")
print(st.get_type())
print(st.get_name())
"#);
    assert_eq!(out, vec!["module", "top"]);
}

#[test]
fn test_symtable_get_symbols_in_scope() {
    let out = run_python(r#"
import symtable
st = symtable.symtable("x = 10; y = 20", "<string>", "exec")
sym_names = [s.get_name() for s in st.get_symbols()]
print("x" in sym_names and "y" in sym_names)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_symtable_symbol_properties_local_vs_global() {
    let out = run_python(r#"
import symtable
code = """
global_var = 100
def func():
    local_var = 50
    return global_var + local_var
"""
st = symtable.symtable(code, "<string>", "exec")
sym_global = st.lookup("global_var")
print(sym_global.is_global())
print(sym_global.is_local())
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_symtable_function_child_symbol_table() {
    let out = run_python(r#"
import symtable
code = """
def my_func(a, b=10):
    c = a + b
    return c
"""
st = symtable.symtable(code, "<string>", "exec")
children = st.get_children()
print(len(children))
child = children[0]
print(child.get_type())
print(child.get_name())
"#);
    assert_eq!(out, vec!["1", "function", "my_func"]);
}

#[test]
fn test_symtable_function_parameters() {
    let out = run_python(r#"
import symtable
code = "def f(arg1, arg2): pass"
st = symtable.symtable(code, "<string>", "exec")
func_st = st.get_children()[0]
p1 = func_st.lookup("arg1")
print(p1.is_parameter())
print(p1.is_local())
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_symtable_imported_symbol() {
    let out = run_python(r#"
import symtable
code = "from math import sqrt, sin"
st = symtable.symtable(code, "<string>", "exec")
s = st.lookup("sqrt")
print(s.is_imported())
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_symtable_class_symbol_table() {
    let out = run_python(r#"
import symtable
code = """
class MyClass:
    attr = 42
    def method(self): pass
"""
st = symtable.symtable(code, "<string>", "exec")
children = st.get_children()
print(len(children))
cls_st = children[0]
print(cls_st.get_type())
print(cls_st.get_name())
"#);
    assert_eq!(out, vec!["1", "class", "MyClass"]);
}

#[test]
fn test_symtable_free_and_cell_variables_nonlocal() {
    let out = run_python(r#"
import symtable
code = """
def outer():
    x = 10
    def inner():
        nonlocal x
        return x
"""
st = symtable.symtable(code, "<string>", "exec")
outer_st = st.get_children()[0]
inner_st = outer_st.get_children()[0]
x_inner = inner_st.lookup("x")
print(x_inner.is_free())
print(x_inner.is_nonlocal())
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_symtable_symbol_table_is_nested() {
    let out = run_python(r#"
import symtable
code = """
def outer():
    def inner(): pass
"""
st = symtable.symtable(code, "<string>", "exec")
outer_st = st.get_children()[0]
inner_st = outer_st.get_children()[0]
print(outer_st.is_nested())
print(inner_st.is_nested())
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_symtable_symbol_is_referenced_and_assigned() {
    let out = run_python(r#"
import symtable
code = "a = 1; b = a"
st = symtable.symtable(code, "<string>", "exec")
sym_a = st.lookup("a")
sym_b = st.lookup("b")
print(sym_a.is_assigned())
print(sym_a.is_referenced())
print(sym_b.is_assigned())
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_symtable_has_children_check() {
    let out = run_python(r#"
import symtable
st1 = symtable.symtable("x = 1", "<string>", "exec")
st2 = symtable.symtable("def f(): pass", "<string>", "exec")
print(st1.has_children())
print(st2.has_children())
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_symtable_get_identifiers_list() {
    let out = run_python(r#"
import symtable
st = symtable.symtable("x = 1; y = 2", "<string>", "exec")
ids = list(st.get_identifiers())
print("x" in ids and "y" in ids)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_symtable_lookup_non_existent_symbol_raises_keyerror() {
    let out = run_python(r#"
import symtable
st = symtable.symtable("x = 1", "<string>", "exec")
try:
    st.lookup("non_existent_symbol")
except KeyError:
    print("KeyError")
"#);
    assert_eq!(out, vec!["KeyError"]);
}

#[test]
fn test_symtable_global_statement_explicit_global() {
    let out = run_python(r#"
import symtable
code = """
x = 1
def f():
    global x
    x = 2
"""
st = symtable.symtable(code, "<string>", "exec")
func_st = st.get_children()[0]
sym_x = func_st.lookup("x")
print(sym_x.is_global())
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_symtable_async_function_symtable() {
    let out = run_python(r#"
import symtable
code = "async def coro(x): await x"
st = symtable.symtable(code, "<string>", "exec")
coro_st = st.get_children()[0]
print(coro_st.is_optimized())
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_symtable_is_optimized_function_scope() {
    let out = run_python(r#"
import symtable
st_mod = symtable.symtable("x = 1", "<string>", "exec")
st_func = symtable.symtable("def f(): x = 1", "<string>", "exec").get_children()[0]
print(st_mod.is_optimized())
print(st_func.is_optimized())
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_symtable_get_lineno_of_symbol_table() {
    let out = run_python(r#"
import symtable
code = "\n\ndef target_func(): pass"
st = symtable.symtable(code, "<string>", "exec")
func_st = st.get_children()[0]
print(func_st.get_lineno())
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_symtable_symbol_is_declared_global() {
    let out = run_python(r#"
import symtable
code = """
def f():
    global g
    g = 10
"""
st = symtable.symtable(code, "<string>", "exec")
func_st = st.get_children()[0]
sym_g = func_st.lookup("g")
print(sym_g.is_declared_global())
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_symtable_annotation_scope_behavior() {
    let out = run_python(r#"
import symtable
code = "x: int = 5"
st = symtable.symtable(code, "<string>", "exec")
sym_x = st.lookup("x")
print(sym_x.is_assigned())
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_symtable_generator_expression_child_scope() {
    let out = run_python(r#"
import symtable
code = "gen = (x * 2 for x in range(10))"
st = symtable.symtable(code, "<string>", "exec")
children = st.get_children()
print(len(children) >= 1)
print(children[0].get_name())
"#);
    assert_eq!(out, vec!["True", "genexpr"]);
}
