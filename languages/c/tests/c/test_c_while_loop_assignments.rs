use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn while_assignment_in_condition() {
    assert_eq!(
        run_c("int main() { int x=3; while(x=x-1) printf(\"%d\", x); return 0; }"),
        vec!["21"]
    );
}
#[test]
fn while_assignment_returned_value() {
    assert_eq!(
        run_c(
            "int f() { static int count=2; return count--; } int main() { int x; while((x=f())) printf(\"%d\", x); return 0; }"
        ),
        vec!["21"]
    );
}
#[test]
fn while_assignment_pointer() {
    assert_eq!(
        run_c(
            "int main() { int arr[] = {1, 2, 0}; int *p = arr; while(*p++) printf(\"%d\", *(p-1)); return 0; }"
        ),
        vec!["12"]
    );
}
#[test]
fn while_assignment_short_circuit() {
    assert_eq!(
        run_c("int main() { int x=0, y=1; while(x && (y=0)) ; printf(\"%d\", y); return 0; }"),
        vec!["1"]
    );
} // y=0 not evaluated
#[test]
fn while_assignment_comma() {
    assert_eq!(
        run_c("int main() { int x=2; while(x--, x) printf(\"%d\", x); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn while_assignment_chained() {
    assert_eq!(
        run_c("int main() { int a, b=2; while(a = b = b - 1) printf(\"%d\", a); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn while_assignment_bitwise() {
    assert_eq!(
        run_c("int main() { int x=3; while(x &= 2) { printf(\"%d\", x); x=0; } return 0; }"),
        vec!["2"]
    );
} // 3 & 2 = 2
#[test]
fn while_assignment_shift() {
    assert_eq!(
        run_c("int main() { int x=4; while(x >>= 1) printf(\"%d\", x); return 0; }"),
        vec!["21"]
    );
}
#[test]
fn while_assignment_compound() {
    assert_eq!(
        run_c("int main() { int x=5; while((x -= 2) > 0) printf(\"%d\", x); return 0; }"),
        vec!["31"]
    );
}
#[test]
fn while_assignment_char() {
    assert_eq!(
        run_c("int main() { char c='B'; while((c = c - 1) >= 'A') printf(\"%c\", c); return 0; }"),
        vec!["A"]
    );
}
#[test]
fn while_assignment_in_func_arg() {
    assert_eq!(
        run_c(
            "int check(int val) { return val; } int main() { int x=2; while(check(x=x-1)) printf(\"%d\", x); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn while_assignment_ternary() {
    assert_eq!(
        run_c("int main() { int x=2; while((x = x>0 ? x-1 : 0)) printf(\"%d\", x); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn while_assignment_struct_member() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S s = {2}; while(s.a--) printf(\"%d\", s.a); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn while_assignment_array_element() {
    assert_eq!(
        run_c("int main() { int arr[1] = {2}; while(arr[0]--) printf(\"%d\", arr[0]); return 0; }"),
        vec!["10"]
    );
}
#[test]
fn while_assignment_implicit_comparison() {
    assert_eq!(
        run_c("int main() { int x=1; while(x=0) printf(\"no\"); printf(\"yes\"); return 0; }"),
        vec!["yes"]
    );
}
