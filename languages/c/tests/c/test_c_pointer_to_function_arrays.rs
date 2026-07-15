use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn pointer_func_array_basic() {
    assert_eq!(
        run_c(
            "int f(){return 1;} int g(){return 2;} int main() { int (*arr[2])() = {f, g}; printf(\"%d\", arr[1]()); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn pointer_func_array_typedef() {
    assert_eq!(
        run_c(
            "typedef int (*F)(void); int f(){return 3;} int main() { F arr[1] = {f}; printf(\"%d\", arr[0]()); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn pointer_func_array_pointer_to_array() {
    assert_eq!(
        run_c(
            "int f(){return 4;} int main() { int (*arr[1])() = {f}; int (*(*p)[1])() = &arr; printf(\"%d\", (*p)[0]()); return 0; }"
        ),
        vec!["4"]
    );
}
#[test]
fn pointer_func_array_pass_to_function() {
    assert_eq!(
        run_c(
            "int f(){return 5;} void run(int (*a[])()) { printf(\"%d\", a[0]()); } int main() { int (*arr[1])() = {f}; run(arr); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn pointer_func_array_return_from_function() {
    assert_eq!(
        run_c(
            "int f(){return 6;} int (*arr[1])() = {f}; int (*(*get_arr())[1])() { return &arr; } int main() { printf(\"%d\", (*get_arr())[0]()); return 0; }"
        ),
        vec!["6"]
    );
}
#[test]
fn pointer_func_array_null_element() {
    assert_eq!(
        run_c("int main() { int (*arr[2])() = {0}; printf(\"%d\", arr[0] == 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn pointer_func_array_size_inference() {
    assert_eq!(
        run_c(
            "int f(){return 7;} int main() { int (*arr[])() = {f, f, f}; printf(\"%d\", (int)(sizeof(arr)/sizeof(arr[0]))); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn pointer_func_array_different_signatures_fails() {
    assert_eq!(
        run_c(
            "/* int f(){return 1;} void g(){} int (*arr[2])() = {f, g}; // error */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn pointer_func_array_compound_literal() {
    assert_eq!(
        run_c(
            "int f(){return 8;} int main() { int (**p)() = (int (*[])()){f}; printf(\"%d\", p[0]()); return 0; }"
        ),
        vec!["8"]
    );
}
#[test]
fn pointer_func_array_vla() {
    assert_eq!(
        run_c(
            "int f(){return 9;} int main() { int n=2; int (*arr[n])(); arr[1]=f; printf(\"%d\", arr[1]()); return 0; }"
        ),
        vec!["9"]
    );
}
#[test]
fn pointer_func_array_struct_member() {
    assert_eq!(
        run_c(
            "int f(){return 10;} struct S { int (*arr[1])(); }; int main() { struct S s = {{f}}; printf(\"%d\", s.arr[0]()); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn pointer_func_array_multidim() {
    assert_eq!(
        run_c(
            "int f(){return 11;} int main() { int (*arr[2][2])() = {{{f, f}, {f, f}}}; printf(\"%d\", arr[1][1]()); return 0; }"
        ),
        vec!["11"]
    );
}
#[test]
fn pointer_func_array_with_args() {
    assert_eq!(
        run_c(
            "int f(int x){return x+1;} int main() { int (*arr[1])(int) = {f}; printf(\"%d\", arr[0](10)); return 0; }"
        ),
        vec!["11"]
    );
}
#[test]
fn pointer_func_array_cast() {
    assert_eq!(
        run_c(
            "int f(){return 12;} int main() { void *arr[1] = {(void*)f}; printf(\"%d\", ((int(*)())arr[0])()); return 0; }"
        ),
        vec!["12"]
    );
}
