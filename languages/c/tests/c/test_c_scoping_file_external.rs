use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn scope_file_global_var() {
    assert_eq!(
        run_c("int g = 10; int main() { printf(\"%d\", g); return 0; }"),
        vec!["10"]
    );
}
#[test]
fn scope_file_static_var() {
    assert_eq!(
        run_c("static int g = 20; int main() { printf(\"%d\", g); return 0; }"),
        vec!["20"]
    );
}
#[test]
fn scope_file_function() {
    assert_eq!(
        run_c("int f() { return 30; } int main() { printf(\"%d\", f()); return 0; }"),
        vec!["30"]
    );
}
#[test]
fn scope_file_static_function() {
    assert_eq!(
        run_c("static int f() { return 40; } int main() { printf(\"%d\", f()); return 0; }"),
        vec!["40"]
    );
}
#[test]
fn scope_file_forward_decl() {
    assert_eq!(
        run_c("extern int g; int main() { printf(\"%d\", g); return 0; } int g = 50;"),
        vec!["50"]
    );
}
#[test]
fn scope_file_forward_func() {
    assert_eq!(
        run_c("int f(void); int main() { printf(\"%d\", f()); return 0; } int f() { return 60; }"),
        vec!["60"]
    );
}
#[test]
fn scope_external_linkage_default() {
    assert_eq!(
        run_c("int g = 70; int main() { extern int g; printf(\"%d\", g); return 0; }"),
        vec!["70"]
    );
}
#[test]
fn scope_external_linkage_func() {
    assert_eq!(
        run_c(
            "int f() { return 80; } int main() { extern int f(); printf(\"%d\", f()); return 0; }"
        ),
        vec!["80"]
    );
}
#[test]
fn scope_local_extern() {
    assert_eq!(
        run_c("int g = 90; int main() { { extern int g; printf(\"%d\", g); } return 0; }"),
        vec!["90"]
    );
}
#[test]
fn scope_local_extern_shadows_local() {
    assert_eq!(
        run_c(
            "int g = 100; int main() { int g = 1; { extern int g; printf(\"%d\", g); } return 0; }"
        ),
        vec!["100"]
    );
}
#[test]
fn scope_multiple_extern_decls() {
    assert_eq!(
        run_c(
            "extern int g; extern int g; int g = 110; int main() { printf(\"%d\", g); return 0; }"
        ),
        vec!["110"]
    );
}
#[test]
fn scope_extern_array_incomplete() {
    assert_eq!(
        run_c(
            "extern int arr[]; int arr[] = {1, 2, 3}; int main() { printf(\"%d\", arr[1]); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn scope_tentative_definition() {
    assert_eq!(
        run_c("int g; int g; int g = 120; int main() { printf(\"%d\", g); return 0; }"),
        vec!["120"]
    );
}
#[test]
fn scope_static_tentative() {
    assert_eq!(
        run_c("static int g; static int g; int main() { g = 130; printf(\"%d\", g); return 0; }"),
        vec!["130"]
    );
}
#[test]
fn scope_file_struct_tag() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S s = {140}; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["140"]
    );
}
