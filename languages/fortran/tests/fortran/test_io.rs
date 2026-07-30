use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: I/O — print, write
// ═══════════════════════════════════════════════════════════

#[test]
fn print_string() {
    let out = run_prints(
        r#"
program test
    print *, "Hello"
end program test
"#,
    );
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn print_integer() {
    let out = run_prints(
        r#"
program test
    print *, 42
end program test
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn print_real() {
    let out = run_prints(
        r#"
program test
    print *, 3.14
end program test
"#,
    );
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn print_multiple() {
    let out = run_prints(
        r#"
program test
    print *, "x =", 42
end program test
"#,
    );
    assert_eq!(out, vec!["x = 42"]);
}

#[test]
fn print_expression() {
    let out = run_prints(
        r#"
program test
    integer :: a, b
    a = 10
    b = 20
    print *, a + b
end program test
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn print_logical() {
    let out = run_prints(
        r#"
program test
    print *, .true.
    print *, .false.
end program test
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn formatted_array_print_ignores_unused_repeat_descriptors() {
    let out = run_prints(
        r#"
program test
    real(8) :: a(2)
    a(1) = 1.25d0
    a(2) = 2.5d0
    print '(4f10.4)', a
end program test
"#,
    );
    assert_eq!(out, vec!["1.2500 2.5000"]);
}

#[test]
fn write_basic() {
    compile_ok(
        r#"
program test
    write(*, *) "Hello from write"
end program test
"#,
    );
}

#[test]
fn print_char_array_join() {
    let out = run_prints(
        r#"
program test
    character(len=5) :: words(2)
    words(1) = "one"
    words(2) = "two"
    print *, words
end program test
"#,
    );
    assert_eq!(out, vec!["one two"]);
}

#[test]
fn write_formatted_integer_width() {
    let out = run_prints(
        r#"
program test
    write(*, '(I4)') 7
end program test
"#,
    );
    assert_eq!(out, vec!["   7"]);
}

#[test]
fn print_logical_and_integer_combo() {
    let out = run_prints(
        r#"
program test
    print *, "alive="
    print *, .true.
    print *, "count="
    print *, 3
end program test
"#,
    );
    assert_eq!(out, vec!["alive=", "true", "count=", "3"]);
}

#[test]
fn print_real_and_character_concat() {
    let out = run_prints(
        r#"
program test
    character(len=8) :: buf
    real :: x
    x = 2.5
    write(buf, '(A,F4.1)') "x=", x
    print *, trim(buf)
end program test
"#,
    );
    assert_eq!(out, vec!["x= 2.5"]);
}

#[test]
fn print_array_comma_separated() {
    let out = run_prints(
        r#"
program test
    integer :: a(3) = [1,2,3]
    print '(3(I0, ","))', a
end program test
"#,
    );
    assert_eq!(out, vec!["1,2,3"]);
}

#[test]
fn write_character_padding_retained() {
    let out = run_prints(
        r#"
program test
    write(*, '(A4)') 'abc'
end program test
"#,
    );
    assert_eq!(out, vec!["abc "]);
}
