use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn generic_selection_int() {
    assert_eq!(
        run_c(
            "int main() { int type_id = _Generic(42, int: 1, float: 2, default: 0); printf(\"%d\", type_id); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn generic_selection_float() {
    assert_eq!(
        run_c(
            "int main() { int type_id = _Generic(3.14f, int: 1, float: 2, default: 0); printf(\"%d\", type_id); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn generic_selection_default() {
    assert_eq!(
        run_c(
            "int main() { int type_id = _Generic(\"hello\", int: 1, float: 2, default: 0); printf(\"%d\", type_id); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn generic_selection_qualifiers() {
    assert_eq!(
        run_c(
            "int main() { const int x = 5; int type_id = _Generic(x, int: 1, default: 0); /* Qualifiers are dropped in expressions */ printf(\"%d\", type_id); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn generic_selection_pointer() {
    assert_eq!(
        run_c(
            "int main() { int x; int *p = &x; int type_id = _Generic(p, int*: 1, int: 2, default: 0); printf(\"%d\", type_id); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn generic_selection_array_decay() {
    assert_eq!(
        run_c(
            "int main() { int arr[5]; int type_id = _Generic(arr, int*: 1, int[5]: 2, default: 0); /* Arrays decay to pointers in _Generic */ printf(\"%d\", type_id); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn generic_selection_function_decay() {
    assert_eq!(
        run_c(
            "void foo() {} int main() { int type_id = _Generic(foo, void(*)(void): 1, default: 0); /* Functions decay to pointers */ printf(\"%d\", type_id); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn generic_selection_macro_wrap() {
    assert_eq!(
        run_c(
            "#define TYPE_NAME(X) _Generic((X), int: \"int\", float: \"float\", default: \"other\")\nint main() { printf(\"%s %s\", TYPE_NAME(5), TYPE_NAME(5.0f)); return 0; }"
        ),
        vec!["int float"]
    );
}
#[test]
fn static_assert_basic() {
    assert_eq!(
        run_c(
            "int main() { _Static_assert(sizeof(int) >= 2, \"int too small\"); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn static_assert_c23_no_message() {
    assert_eq!(
        run_c("int main() { _Static_assert(1 == 1); printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
#[test]
fn static_assert_global_scope() {
    assert_eq!(
        run_c(
            "_Static_assert(sizeof(char) == 1, \"char size\");\nint main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn static_assert_struct_scope() {
    assert_eq!(
        run_c(
            "struct Foo { int x; _Static_assert(sizeof(int) == 4, \"\"); };\nint main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
