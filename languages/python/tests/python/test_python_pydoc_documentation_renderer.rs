use super::helpers::run_python;

// pydoc — locate, render_doc, TextDoc, HTMLDoc, ispackage, synopsis, help, describe

#[test]
fn test_pydoc_locate_module_or_object() {
    let out = run_python(r#"
import pydoc, math
obj = pydoc.locate("math.sin")
print(obj is math.sin)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pydoc_render_doc_text_mode() {
    let out = run_python(r#"
import pydoc

def sample_func(a, b):
    """Sample docstring."""
    return a + b

doc_str = pydoc.render_doc(sample_func)
print("sample_func" in doc_str)
print("Sample docstring." in doc_str)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_pydoc_text_doc_renderer_class() {
    let out = run_python(r#"
import pydoc

class Target:
    """Class docstring."""

text_doc = pydoc.TextDoc()
doc = text_doc.docclass(Target)
print("Target" in doc)
print("Class docstring." in doc)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_pydoc_html_doc_renderer_class() {
    let out = run_python(r#"
import pydoc

def dummy(): pass

html_doc = pydoc.HTMLDoc()
doc = html_doc.docroutine(dummy)
print("<a href=" in doc or "dummy" in doc)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pydoc_ispackage_check() {
    let out = run_python(r#"
import pydoc, os, json
print(pydoc.ispackage("json"))
print(pydoc.ispackage("os"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_pydoc_synopsis_module_summary() {
    let out = run_python(r#"
import pydoc, json
syn = pydoc.synopsis(json.__file__)
print(isinstance(syn, str) or syn is None)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pydoc_describe_object_type_name() {
    let out = run_python(r#"
import pydoc

def fn(): pass
class C: pass

print(pydoc.describe(fn))
print(pydoc.describe(C))
print(pydoc.describe(123))
"#);
    assert_eq!(out, vec!["function fn", "class C", "int object"]);
}

#[test]
fn test_pydoc_text_doc_docfunction() {
    let out = run_python(r#"
import pydoc

def calc(x, y=10):
    """Calculates something."""
    return x * y

td = pydoc.TextDoc()
res = td.docfunction(calc)
print("calc(x, y=10)" in res or "calc" in res)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pydoc_text_doc_docmodule() {
    let out = run_python(r#"
import pydoc, math
td = pydoc.TextDoc()
doc = td.docmodule(math)
print("NAME" in doc)
print("math" in doc)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_pydoc_plain_formatting() {
    let out = run_python(r#"
import pydoc
text = "\x1b[1mBold Text\x1b[0m"
plain = pydoc.plain(text)
print(plain)
"#);
    assert_eq!(out, vec!["Bold Text"]);
}

#[test]
fn test_pydoc_helper_object_instance() {
    let out = run_python(r#"
import pydoc
h = pydoc.Helper()
print(hasattr(h, "help"))
print(hasattr(h, "intro"))
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_pydoc_stripid_string_transformation() {
    let out = run_python(r#"
import pydoc
s = "<object at 0x7f9a1b2c3d4e>"
stripped = pydoc.stripid(s)
print(stripped)
"#);
    assert_eq!(out, vec!["<object at 0x...>"]);
}

#[test]
fn test_pydoc_resolve_symbol_path() {
    let out = run_python(r#"
import pydoc, math
obj, name = pydoc.resolve("math.sqrt")
print(obj is math.sqrt)
print(name)
"#);
    assert_eq!(out, vec!["True", "math.sqrt"]);
}

#[test]
fn test_pydoc_classname_formatting() {
    let out = run_python(r#"
import pydoc, json

class Sub(json.JSONDecoder): pass

name = pydoc.classname(Sub, "json")
print(name)
"#);
    assert_eq!(out, vec!["Sub"]);
}

#[test]
fn test_pydoc_isdata_type_check() {
    let out = run_python(r#"
import pydoc

def fn(): pass
x = [1, 2, 3]

print(pydoc.isdata(fn))
print(pydoc.isdata(x))
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_pydoc_text_doc_docother() {
    let out = run_python(r#"
import pydoc
td = pydoc.TextDoc()
doc = td.docother(42, "MY_CONST")
print("MY_CONST = 42" in doc)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pydoc_allmethods_class_inspection() {
    let out = run_python(r#"
import pydoc

class A:
    def method_a(self): pass

class B(A):
    def method_b(self): pass

methods = pydoc.allmethods(B)
names = [m.__name__ for m in methods]
print("method_a" in names and "method_b" in names)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pydoc_splitdoc_docstring_header() {
    let out = run_python(r#"
import pydoc
doc = "First line summary.\n\nDetailed explanation line 2."
head, tail = pydoc.splitdoc(doc)
print(head)
print(tail)
"#);
    assert_eq!(out, vec!["First line summary.", "Detailed explanation line 2."]);
}

#[test]
fn test_pydoc_locate_error_handling() {
    let out = run_python(r#"
import pydoc
try:
    obj = pydoc.locate("non_existent_module_9999.invalid")
    print(obj is None)
except Exception:
    print(True)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pydoc_render_doc_renderer_choice() {
    let out = run_python(r#"
import pydoc, math
doc_text = pydoc.render_doc(math, renderer=pydoc.plaintext)
print("math" in doc_text)
"#);
    assert_eq!(out, vec!["True"]);
}
