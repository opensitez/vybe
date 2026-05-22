use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: I/O — print, write
// ═══════════════════════════════════════════════════════════

#[test]
fn print_string() {
    let out = run_prints(r#"
program test
    print *, "Hello"
end program test
"#);
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn print_integer() {
    let out = run_prints(r#"
program test
    print *, 42
end program test
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn print_real() {
    let out = run_prints(r#"
program test
    print *, 3.14
end program test
"#);
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn print_multiple() {
    let out = run_prints(r#"
program test
    print *, "x =", 42
end program test
"#);
    assert_eq!(out, vec!["x = 42"]);
}

#[test]
fn print_expression() {
    let out = run_prints(r#"
program test
    integer :: a, b
    a = 10
    b = 20
    print *, a + b
end program test
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn print_logical() {
    let out = run_prints(r#"
program test
    print *, .true.
    print *, .false.
end program test
"#);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn formatted_array_print_ignores_unused_repeat_descriptors() {
    let out = run_prints(r#"
program test
    real(8) :: a(2)
    a(1) = 1.25d0
    a(2) = 2.5d0
    print '(4f10.4)', a
end program test
"#);
    assert_eq!(out, vec!["1.2500 2.5000"]);
}

#[test]
fn write_basic() {
    compile_ok(r#"
program test
    write(*, *) "Hello from write"
end program test
"#);
}
