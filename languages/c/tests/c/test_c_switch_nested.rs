use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn switch_nested_basic() {
    assert_eq!(
        run_c(
            "int main() { int x=1, y=2; switch(x) { case 1: switch(y) { case 2: printf(\"A\"); break; } break; } return 0; }"
        ),
        vec!["A"]
    );
}
#[test]
fn switch_nested_default() {
    assert_eq!(
        run_c(
            "int main() { int x=1, y=5; switch(x) { case 1: switch(y) { default: printf(\"D\"); break; } break; } return 0; }"
        ),
        vec!["D"]
    );
}
#[test]
fn switch_nested_break_scope() {
    assert_eq!(
        run_c(
            "int main() { int x=1, y=2; switch(x) { case 1: switch(y) { case 2: break; } printf(\"InnerEnd\"); break; } return 0; }"
        ),
        vec!["InnerEnd"]
    );
} // break only exits inner
#[test]
fn switch_nested_fallthrough() {
    assert_eq!(
        run_c(
            "int main() { int x=1, y=2; switch(x) { case 1: switch(y) { case 2: printf(\"2\"); } case 2: printf(\"Out2\"); break; } return 0; }"
        ),
        vec!["2Out2"]
    );
} // Fallthrough from inner switch block to outer case
#[test]
fn switch_nested_same_case_values() {
    assert_eq!(
        run_c(
            "int main() { int x=1, y=1; switch(x) { case 1: switch(y) { case 1: printf(\"Inner\"); break; } break; } return 0; }"
        ),
        vec!["Inner"]
    );
}
#[test]
fn switch_nested_goto_cross() {
    assert_eq!(
        run_c(
            "int main() { int x=1; switch(x) { case 1: goto L; case 2: switch(2) { case 2: L: printf(\"Jumped\"); break; } break; } return 0; }"
        ),
        vec!["Jumped"]
    );
}
#[test]
fn switch_nested_deep() {
    assert_eq!(
        run_c(
            "int main() { switch(1) { case 1: switch(2) { case 2: switch(3) { case 3: printf(\"3\"); break; } break; } break; } return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn switch_nested_shadowed_vars() {
    assert_eq!(
        run_c(
            "int main() { switch(1) { case 1: { int a = 5; switch(2) { case 2: { int a = 10; printf(\"%d\", a); break; } } break; } } return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn switch_nested_with_while() {
    assert_eq!(
        run_c(
            "int main() { switch(1) { case 1: while(1) { switch(2) { case 2: printf(\"W\"); break; } break; } break; } return 0; }"
        ),
        vec!["W"]
    );
}
#[test]
fn switch_nested_case_belongs_to_inner() {
    assert_eq!(
        run_c(
            "int main() { int x=1, y=2; switch(x) { case 1: switch(y) { case 2: printf(\"Inner\"); break; case 3: printf(\"Inner3\"); break; } break; } return 0; }"
        ),
        vec!["Inner"]
    );
}
#[test]
fn switch_nested_inner_default_outer_case() {
    assert_eq!(
        run_c(
            "int main() { int x=1, y=5; switch(x) { case 1: switch(y) { default: printf(\"D\"); break; } break; case 5: printf(\"O\"); break; } return 0; }"
        ),
        vec!["D"]
    );
}
#[test]
fn switch_nested_outer_default_inner_case() {
    assert_eq!(
        run_c(
            "int main() { int x=5, y=1; switch(x) { default: switch(y) { case 1: printf(\"I\"); break; } break; } return 0; }"
        ),
        vec!["I"]
    );
}
#[test]
fn switch_nested_unreachable_inner() {
    assert_eq!(
        run_c(
            "int main() { switch(1) { case 1: switch(2) { printf(\"U\"); case 2: printf(\"M\"); break; } break; } return 0; }"
        ),
        vec!["M"]
    );
}
#[test]
fn switch_nested_continue_in_loop() {
    assert_eq!(
        run_c(
            "int main() { for(int i=0; i<2; i++) { switch(i) { case 0: switch(1) { case 1: continue; } break; case 1: printf(\"1\"); break; } } return 0; }"
        ),
        vec!["1"]
    );
} // continue works in switch if inside loop
#[test]
fn switch_nested_break_to_outer_loop() {
    assert_eq!(
        run_c(
            "int main() { while(1) { switch(1) { case 1: switch(2) { case 2: break; } break; } printf(\"End\"); break; } return 0; }"
        ),
        vec!["End"]
    );
} // break only exits switch, loop continues, then second break exits loop
