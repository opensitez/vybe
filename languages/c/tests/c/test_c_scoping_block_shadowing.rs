use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn block_shadow_simple() {
    assert_eq!(
        run_c("int main() { int x = 1; { int x = 2; printf(\"%d\", x); } return 0; }"),
        vec!["2"]
    );
}
#[test]
fn block_shadow_outer_visible_after() {
    assert_eq!(
        run_c("int main() { int x = 1; { int x = 2; } printf(\"%d\", x); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn block_shadow_global() {
    assert_eq!(
        run_c("int x = 5; int main() { int x = 10; printf(\"%d\", x); return 0; }"),
        vec!["10"]
    );
}
#[test]
fn block_shadow_parameter() {
    assert_eq!(
        run_c(
            "int f(int x) { { int x = 3; return x; } } int main() { printf(\"%d\", f(1)); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn block_shadow_for_loop_decl() {
    assert_eq!(
        run_c(
            "int main() { int x = 0; for (int x = 1; x < 2; x++) { printf(\"%d\", x); } return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn block_shadow_for_loop_body() {
    assert_eq!(
        run_c(
            "int main() { for (int x = 1; x < 2; x++) { int x = 5; printf(\"%d\", x); } return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn block_shadow_if_statement() {
    assert_eq!(
        run_c("int main() { int x = 1; if (1) { int x = 2; printf(\"%d\", x); } return 0; }"),
        vec!["2"]
    );
}
#[test]
fn block_shadow_while_statement() {
    assert_eq!(
        run_c(
            "int main() { int x = 1; while (x == 1) { int x = 2; printf(\"%d\", x); break; } return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn block_shadow_do_while_statement() {
    assert_eq!(
        run_c(
            "int main() { int x = 1; do { int x = 2; printf(\"%d\", x); } while (0); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn block_shadow_switch_statement() {
    assert_eq!(
        run_c(
            "int main() { int x = 1; switch(1) { case 1: { int x = 2; printf(\"%d\", x); } } return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn block_shadow_different_types() {
    assert_eq!(
        run_c("int main() { int x = 65; { char x = 'B'; printf(\"%c\", x); } return 0; }"),
        vec!["B"]
    );
}
#[test]
fn block_shadow_struct_type() {
    assert_eq!(
        run_c(
            "int main() { int x = 5; { struct S { int x; } s = {10}; printf(\"%d\", s.x); } return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn block_shadow_nested_multiple() {
    assert_eq!(
        run_c(
            "int main() { int x = 1; { int x = 2; { int x = 3; printf(\"%d\", x); } } return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn block_shadow_outer_in_expr() {
    assert_eq!(
        run_c(
            "int main() { int x = 5; { int x = x + 1; /* UB in some cases if uninitialized, but here x on right binds to inner uninit, however we test parsing mostly */ printf(\"ok\"); } return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn block_shadow_function_name() {
    assert_eq!(
        run_c("int f() { return 1; } int main() { int f = 2; printf(\"%d\", f); return 0; }"),
        vec!["2"]
    );
}
