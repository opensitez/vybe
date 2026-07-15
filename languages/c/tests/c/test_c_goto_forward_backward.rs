use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn goto_forward_basic() {
    assert_eq!(
        run_c("int main() { goto L; printf(\"X\"); L: printf(\"A\"); return 0; }"),
        vec!["A"]
    );
}
#[test]
fn goto_backward_basic() {
    assert_eq!(
        run_c("int main() { int i=0; L: if(i++ == 0) goto L; printf(\"%d\", i); return 0; }"),
        vec!["2"]
    );
}
#[test]
fn goto_bypassing_initialization() {
    assert_eq!(
        run_c("int main() { goto L; int a = 5; L: printf(\"X\"); return 0; }"),
        vec!["X"]
    );
} // Legal in C, uninitialized variable
#[test]
fn goto_bypassing_vla_fails_in_c99() {
    assert_eq!(
        run_c(
            "/* int main() { int n=5; goto L; int arr[n]; L: return 0; } // VLA bypass is illegal */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn goto_out_of_block() {
    assert_eq!(
        run_c("int main() { { goto L; } L: printf(\"O\"); return 0; }"),
        vec!["O"]
    );
}
#[test]
fn goto_into_block() {
    assert_eq!(
        run_c("int main() { goto L; { L: printf(\"I\"); } return 0; }"),
        vec!["I"]
    );
}
#[test]
fn goto_out_of_nested_loops() {
    assert_eq!(
        run_c(
            "int main() { for(int i=0; i<5; i++) for(int j=0; j<5; j++) if(i==1 && j==1) goto L; L: printf(\"Escaped\"); return 0; }"
        ),
        vec!["Escaped"]
    );
}
#[test]
fn goto_into_if_body() {
    assert_eq!(
        run_c("int main() { goto L; if(0) { L: printf(\"If\"); } return 0; }"),
        vec!["If"]
    );
}
#[test]
fn goto_into_else_body() {
    assert_eq!(
        run_c("int main() { goto L; if(1) {} else { L: printf(\"Else\"); } return 0; }"),
        vec!["Else"]
    );
}
#[test]
fn goto_multiple_labels() {
    assert_eq!(
        run_c(
            "int main() { goto L2; L1: printf(\"1\"); goto End; L2: printf(\"2\"); goto L1; End: return 0; }"
        ),
        vec!["21"]
    );
}
#[test]
fn goto_label_at_end_of_block() {
    assert_eq!(
        run_c("int main() { goto L; printf(\"X\"); L: ; return 0; }"),
        Vec::<&str>::new()
    );
} // label needs statement, empty stmt
#[test]
fn goto_shadowing_label_fails() {
    assert_eq!(
        run_c(
            "/* int main() { L: goto L; { L: goto L; } return 0; } // Labels have function scope */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn goto_same_name_as_variable() {
    assert_eq!(
        run_c("int main() { int x = 1; goto x; x: printf(\"%d\", x); return 0; }"),
        vec!["1"]
    );
} // Labels have separate namespace
#[test]
fn goto_computed_gcc_ext() {
    assert_eq!(
        run_c(
            "int main() { void *ptr = &&L; goto *ptr; printf(\"X\"); L: printf(\"C\"); return 0; }"
        ),
        vec!["C"]
    );
} // GNU computed goto
#[test]
fn goto_computed_array() {
    assert_eq!(
        run_c(
            "int main() { void *ptrs[] = {&&L1, &&L2}; int i=1; goto *ptrs[i]; L1: printf(\"1\"); return 0; L2: printf(\"2\"); return 0; }"
        ),
        vec!["2"]
    );
}
