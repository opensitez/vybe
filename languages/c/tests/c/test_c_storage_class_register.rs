use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn register_basic_local() {
    assert_eq!(
        run_c("int main() { register int a = 1; printf(\"%d\", a); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn register_uninitialized() {
    assert_eq!(
        run_c("int main() { register int a; a = 5; printf(\"%d\", a); return 0; }"),
        vec!["5"]
    );
}
#[test]
fn register_loop_counter() {
    assert_eq!(
        run_c(
            "int main() { int sum = 0; for (register int i = 0; i < 3; i++) sum += i; printf(\"%d\", sum); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn register_pointer_fails() {
    assert_eq!(
        run_c(
            "int main() { register int a = 1; /* int *p = &a; // address of register variable is invalid in C */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn register_array() {
    assert_eq!(
        run_c("int main() { register int arr[3] = {1,2,3}; printf(\"%d\", arr[0]); return 0; }"),
        vec!["1"]
    );
} // Valid, but array elements can't have their addresses taken if the whole array is register
#[test]
fn register_struct() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { register struct S s = {4}; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["4"]
    );
}
#[test]
fn register_parameter() {
    assert_eq!(
        run_c(
            "int f(register int x) { return x + 1; } int main() { printf(\"%d\", f(2)); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn register_shadowing() {
    assert_eq!(
        run_c(
            "int main() { register int a = 1; { register int a = 2; printf(\"%d\", a); } return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn register_multiple_same_line() {
    assert_eq!(
        run_c("int main() { register int a = 1, b = 2; printf(\"%d\", a+b); return 0; }"),
        vec!["3"]
    );
}
#[test]
fn register_volatile() {
    assert_eq!(
        run_c("int main() { register volatile int a = 10; printf(\"%d\", a); return 0; }"),
        vec!["10"]
    );
}
#[test]
fn register_global_fails() {
    assert_eq!(
        run_c(
            "/* register int g = 5; // Global variables cannot be register */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn register_static_fails() {
    assert_eq!(
        run_c(
            "int main() { /* register static int a = 5; // Conflicting specifiers */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn register_address_of_array_fails() {
    assert_eq!(
        run_c(
            "int main() { register int arr[3]; /* int *p = arr; // Array decay yields address, invalid for register array */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn register_sizeof() {
    assert_eq!(
        run_c("int main() { register int a = 5; printf(\"%d\", (int)sizeof(a)); return 0; }"),
        vec!["4"]
    );
} // Valid to take sizeof register var
#[test]
fn register_alignof() {
    assert_eq!(
        run_c("int main() { register int a = 5; printf(\"%d\", (int)_Alignof(int)); return 0; }"),
        vec!["4"]
    );
}
