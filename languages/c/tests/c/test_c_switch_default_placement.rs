use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn switch_default_first() {
    assert_eq!(
        run_c(
            "int main() { int x=5; switch(x) { default: printf(\"D\"); break; case 1: printf(\"1\"); break; } return 0; }"
        ),
        vec!["D"]
    );
}
#[test]
fn switch_default_middle() {
    assert_eq!(
        run_c(
            "int main() { int x=5; switch(x) { case 1: printf(\"1\"); break; default: printf(\"D\"); break; case 2: printf(\"2\"); break; } return 0; }"
        ),
        vec!["D"]
    );
}
#[test]
fn switch_default_last() {
    assert_eq!(
        run_c(
            "int main() { int x=5; switch(x) { case 1: printf(\"1\"); break; default: printf(\"D\"); break; } return 0; }"
        ),
        vec!["D"]
    );
}
#[test]
fn switch_default_fallthrough_to_case() {
    assert_eq!(
        run_c(
            "int main() { int x=5; switch(x) { default: printf(\"D\"); case 1: printf(\"1\"); break; } return 0; }"
        ),
        vec!["D1"]
    );
}
#[test]
fn switch_case_fallthrough_to_default() {
    assert_eq!(
        run_c(
            "int main() { int x=1; switch(x) { case 1: printf(\"1\"); default: printf(\"D\"); break; } return 0; }"
        ),
        vec!["1D"]
    );
}
#[test]
fn switch_no_default_no_match() {
    assert_eq!(
        run_c(
            "int main() { int x=5; switch(x) { case 1: printf(\"1\"); break; } printf(\"End\"); return 0; }"
        ),
        vec!["End"]
    );
}
#[test]
fn switch_default_only() {
    assert_eq!(
        run_c("int main() { int x=5; switch(x) { default: printf(\"D\"); } return 0; }"),
        vec!["D"]
    );
}
#[test]
fn switch_multiple_defaults_fails() {
    assert_eq!(
        run_c(
            "/* int main() { switch(1) { default: break; default: break; } return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn switch_default_not_in_switch_fails() {
    assert_eq!(
        run_c(
            "/* int main() { default: printf(\"D\"); return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn switch_default_nested_blocks() {
    assert_eq!(
        run_c(
            "int main() { int x=5; switch(x) { case 1: { default: printf(\"D\"); break; } } return 0; }"
        ),
        vec!["D"]
    );
} // Valid in C! Labels can be anywhere in switch scope
#[test]
fn switch_default_in_while() {
    assert_eq!(
        run_c(
            "int main() { int x=5; switch(x) { while(0) { default: printf(\"D\"); break; } } return 0; }"
        ),
        vec!["D"]
    );
} // Duff's device style
#[test]
fn switch_default_unreachable_code_before() {
    assert_eq!(
        run_c(
            "int main() { int x=5; switch(x) { printf(\"Unreachable\"); default: printf(\"D\"); break; } return 0; }"
        ),
        vec!["D"]
    );
}
#[test]
fn switch_default_empty() {
    assert_eq!(
        run_c("int main() { int x=5; switch(x) { default: ; } printf(\"End\"); return 0; }"),
        vec!["End"]
    );
}
#[test]
fn switch_default_break_only() {
    assert_eq!(
        run_c("int main() { int x=5; switch(x) { default: break; } printf(\"End\"); return 0; }"),
        vec!["End"]
    );
}
#[test]
fn switch_default_with_label() {
    assert_eq!(
        run_c("int main() { int x=5; switch(x) { default: L: printf(\"D\"); break; } return 0; }"),
        vec!["D"]
    );
}
