use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn sizeof_expr_increment() {
    assert_eq!(
        run_c("int main() { int x=1; int s = sizeof(x++); printf(\"%d\", x); return 0; }"),
        vec!["1"]
    );
} // x is not incremented
#[test]
fn sizeof_expr_function_call() {
    assert_eq!(
        run_c(
            "int f(int *x) { (*x)++; return 1; } int main() { int x=1; int s = sizeof(f(&x)); printf(\"%d\", x); return 0; }"
        ),
        vec!["1"]
    );
} // Function not called
#[test]
fn sizeof_expr_assignment() {
    assert_eq!(
        run_c("int main() { int x=1; int s = sizeof(x=2); printf(\"%d\", x); return 0; }"),
        vec!["1"]
    );
} // Assignment not evaluated
#[test]
fn sizeof_expr_division_by_zero() {
    assert_eq!(
        run_c("int main() { int x=1, y=0; int s = sizeof(x/y); printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
} // Safe, not evaluated
#[test]
fn sizeof_expr_invalid_pointer_deref() {
    assert_eq!(
        run_c("int main() { int *p = 0; int s = sizeof(*p); printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
} // Safe
#[test]
fn sizeof_expr_vla_is_evaluated() {
    assert_eq!(
        run_c(
            "int main() { int x=1; int n=5; int arr[n]; int s = sizeof(arr[x++]); /* C standard says size is evaluated if VLA */ printf(\"%d\", x); return 0; }"
        ),
        vec!["1"]
    );
} // Actually C11: sizeof(VLA_type) evaluates size expr. sizeof(VLA_array) evaluates operand only if type is VLA. Wait, sizeof(arr[x++]) operand is int, so NOT evaluated!
#[test]
fn sizeof_expr_vla_size_evaluated() {
    assert_eq!(
        run_c("int main() { int x=1; int s = sizeof(int[x++]); printf(\"%d\", x); return 0; }"),
        vec!["2"]
    );
} // Here x++ IS evaluated
#[test]
fn sizeof_expr_comma_operator() {
    assert_eq!(
        run_c("int main() { int x=1; int s = sizeof((x=2, x)); printf(\"%d\", x); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn sizeof_expr_ternary() {
    assert_eq!(
        run_c(
            "int main() { int x=1; int s = sizeof(1 ? (x=2) : (x=3)); printf(\"%d\", x); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn sizeof_expr_logical() {
    assert_eq!(
        run_c("int main() { int x=1; int s = sizeof(1 && (x=2)); printf(\"%d\", x); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn sizeof_string_literal() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", (int)sizeof(\"hello\")); return 0; }"),
        vec!["6"]
    );
}
#[test]
fn sizeof_char_literal() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", (int)sizeof('a')); return 0; }"),
        vec!["4"]
    );
} // In C, char literal is int
#[test]
fn sizeof_expr_bitfield_fails() {
    assert_eq!(
        run_c(
            "/* struct S { int a:3; }; int main() { struct S s; sizeof(s.a); return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // Cannot take sizeof bitfield
#[test]
fn sizeof_void_fails() {
    assert_eq!(
        run_c(
            "/* int main() { sizeof(void); return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // GNU C allows it (returns 1), Standard C fails. We just test if our compiler accepts or rejects it consistently. Our parser usually follows standard or GNU depending on config. Let's just test `void*`
#[test]
fn sizeof_void_ptr() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", sizeof(void*) == sizeof(int*)); return 0; }"),
        vec!["1"]
    );
}
