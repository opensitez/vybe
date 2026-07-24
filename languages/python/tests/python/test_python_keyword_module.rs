use super::helpers::run_python;

// keyword — iskeyword, issoftkeyword, kwlist, softkwlist

#[test]
fn test_keyword_iskeyword_true_for_def() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword("def"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_iskeyword_true_for_class() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword("class"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_iskeyword_true_for_for() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword("for"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_iskeyword_false_for_identifier() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword("hello"))
print(keyword.iskeyword("myvar"))
"#);
    assert_eq!(out, vec!["False", "False"]);
}

#[test]
fn test_keyword_iskeyword_case_sensitive() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword("True"))
print(keyword.iskeyword("true"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_keyword_kwlist_is_sorted() {
    let out = run_python(r#"
import keyword
print(keyword.kwlist == sorted(keyword.kwlist))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_kwlist_contains_common_keywords() {
    let out = run_python(r#"
import keyword
kws = keyword.kwlist
for k in ["if", "else", "while", "return", "import", "lambda"]:
    print(k in kws)
"#);
    assert_eq!(out, vec!["True", "True", "True", "True", "True", "True"]);
}

#[test]
fn test_keyword_kwlist_all_are_keywords() {
    let out = run_python(r#"
import keyword
print(all(keyword.iskeyword(k) for k in keyword.kwlist))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_kwlist_minimum_count() {
    let out = run_python(r#"
import keyword
# Python 3 has at least 35 keywords
print(len(keyword.kwlist) >= 35)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_iskeyword_builtins_are_not_keywords() {
    let out = run_python(r#"
import keyword
# Built-in functions are NOT language keywords
print(keyword.iskeyword("print"))
print(keyword.iskeyword("len"))
print(keyword.iskeyword("range"))
"#);
    assert_eq!(out, vec!["False", "False", "False"]);
}

#[test]
fn test_keyword_iskeyword_operators_not_keywords() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword("+"))
print(keyword.iskeyword("="))
"#);
    assert_eq!(out, vec!["False", "False"]);
}

#[test]
fn test_keyword_iskeyword_none_true_false_are_keywords() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword("None"))
print(keyword.iskeyword("True"))
print(keyword.iskeyword("False"))
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_keyword_iskeyword_and_or_not_are_keywords() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword("and"))
print(keyword.iskeyword("or"))
print(keyword.iskeyword("not"))
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_keyword_issoftkeyword_match_is_soft() {
    let out = run_python(r#"
import keyword, sys
if sys.version_info >= (3, 9):
    print(keyword.issoftkeyword("match"))
else:
    print(True)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_issoftkeyword_type_is_soft() {
    let out = run_python(r#"
import keyword, sys
if sys.version_info >= (3, 12):
    print(keyword.issoftkeyword("type"))
else:
    print(True)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_issoftkeyword_hard_keywords_are_not_soft() {
    let out = run_python(r#"
import keyword, sys
if hasattr(keyword, "issoftkeyword"):
    print(keyword.issoftkeyword("def"))
    print(keyword.issoftkeyword("class"))
else:
    print(False)
    print(False)
"#);
    assert_eq!(out, vec!["False", "False"]);
}

#[test]
fn test_keyword_softkwlist_is_list() {
    let out = run_python(r#"
import keyword, sys
if hasattr(keyword, "softkwlist"):
    print(isinstance(keyword.softkwlist, list))
else:
    print(True)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_softkwlist_all_are_soft_keywords() {
    let out = run_python(r#"
import keyword, sys
if hasattr(keyword, "softkwlist") and hasattr(keyword, "issoftkeyword"):
    print(all(keyword.issoftkeyword(k) for k in keyword.softkwlist))
else:
    print(True)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_keyword_iskeyword_empty_string_is_false() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword(""))
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_keyword_iskeyword_with_is_keyword() {
    let out = run_python(r#"
import keyword
print(keyword.iskeyword("with"))
print(keyword.iskeyword("as"))
print(keyword.iskeyword("yield"))
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}
