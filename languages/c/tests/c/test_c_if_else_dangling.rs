use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn if_else_dangling_basic() {
    assert_eq!(
        run_c("int main() { int x=1; if (1) if (0) x=2; else x=3; printf(\"%d\", x); return 0; }"),
        vec!["3"]
    );
} // else binds to inner
#[test]
fn if_else_dangling_braces() {
    assert_eq!(
        run_c(
            "int main() { int x=1; if (1) { if (0) x=2; } else x=3; printf(\"%d\", x); return 0; }"
        ),
        vec!["1"]
    );
} // braces force else to outer
#[test]
fn if_else_dangling_nested_else() {
    assert_eq!(
        run_c(
            "int main() { int x=1; if (1) if (1) if (0) x=2; else x=3; else x=4; else x=5; printf(\"%d\", x); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn if_else_dangling_while() {
    assert_eq!(
        run_c(
            "int main() { int x=1; if (1) while(0) if (1) x=2; else x=3; printf(\"%d\", x); return 0; }"
        ),
        vec!["1"]
    );
} // else binds to if inside while
#[test]
fn if_else_dangling_for() {
    assert_eq!(
        run_c(
            "int main() { int x=1; if (1) for(;;) if (1) { x=2; break; } else x=3; printf(\"%d\", x); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn if_else_dangling_macro() {
    assert_eq!(
        run_c(
            "#define IF_TRUE if (1)\nint main() { int x=1; IF_TRUE if (0) x=2; else x=3; printf(\"%d\", x); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn if_else_dangling_macro_with_else() {
    assert_eq!(
        run_c(
            "#define COND if (0) {} else\nint main() { int x=1; COND if (1) x=2; else x=3; printf(\"%d\", x); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn if_else_dangling_empty_then() {
    assert_eq!(
        run_c("int main() { int x=1; if (1) if (0) ; else x=3; printf(\"%d\", x); return 0; }"),
        vec!["3"]
    );
}
#[test]
fn if_else_dangling_empty_else() {
    assert_eq!(
        run_c("int main() { int x=1; if (1) if (0) x=2; else; printf(\"%d\", x); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn if_else_dangling_switch() {
    assert_eq!(
        run_c(
            "int main() { int x=1; if (1) switch(0) case 0: if (1) x=2; else x=3; printf(\"%d\", x); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn if_else_dangling_label() {
    assert_eq!(
        run_c(
            "int main() { int x=1; if (1) L: if (0) x=2; else x=3; printf(\"%d\", x); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn if_else_dangling_goto() {
    assert_eq!(
        run_c(
            "int main() { int x=1; if (1) if (0) goto L; else x=3; L: printf(\"%d\", x); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn if_else_chained() {
    assert_eq!(
        run_c(
            "int main() { int x=0; if (0) x=1; else if (0) x=2; else if (1) x=3; else x=4; printf(\"%d\", x); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn if_else_dangling_compound_stmt() {
    assert_eq!(
        run_c(
            "int main() { int x=1; if (1) { if (0) x=2; else x=3; } printf(\"%d\", x); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn if_else_dangling_do_while() {
    assert_eq!(
        run_c(
            "int main() { int x=1; if (1) do if (0) x=2; else x=3; while(0); printf(\"%d\", x); return 0; }"
        ),
        vec!["3"]
    );
}
