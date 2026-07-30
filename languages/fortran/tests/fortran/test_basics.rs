use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: Basic programs, variables, assignment
// ═══════════════════════════════════════════════════════════

#[test]
fn hello_world() {
    let out = run_prints(
        r#"
program hello
    print *, "Hello, World!"
end program hello
"#,
    );
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn integer_variable() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 42
    print *, x
end program test
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn real_variable() {
    let out = run_prints(
        r#"
program test
    real :: pi
    pi = 3.14159
    print *, pi
end program test
"#,
    );
    assert_eq!(out, vec!["3.14159"]);
}

#[test]
fn character_variable() {
    let out = run_prints(
        r#"
program test
    character(len=20) :: name
    name = "Fortran"
    print *, name
end program test
"#,
    );
    assert_eq!(out, vec!["Fortran"]);
}

#[test]
fn logical_variable() {
    let out = run_prints(
        r#"
program test
    logical :: flag
    flag = .true.
    print *, flag
end program test
"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn variable_initialization() {
    let out = run_prints(
        r#"
program test
    integer :: x = 10
    integer :: y = 20
    print *, x + y
end program test
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn arithmetic_basic() {
    let out = run_prints(
        r#"
program test
    integer :: a, b
    a = 10
    b = 3
    print *, a + b
    print *, a - b
    print *, a * b
end program test
"#,
    );
    assert_eq!(out, vec!["13", "7", "30"]);
}

#[test]
fn exponentiation() {
    let out = run_prints(
        r#"
program test
    print *, 2 ** 10
end program test
"#,
    );
    assert_eq!(out, vec!["1024"]);
}

#[test]
fn string_concat() {
    let out = run_prints(
        r#"
program test
    character(len=20) :: greeting
    greeting = "Hello" // " " // "World"
    print *, greeting
end program test
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn implicit_none() {
    let out = run_prints(
        r#"
program test
    implicit none
    integer :: x
    x = 42
    print *, x
end program test
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn multiple_declarations() {
    let out = run_prints(
        r#"
program test
    integer :: a = 1, b = 2, c = 3
    print *, a + b + c
end program test
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn parameter_constant() {
    let out = run_prints(
        r#"
program test
    integer, parameter :: MAX_SIZE = 100
    print *, MAX_SIZE
end program test
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn implicit_none_multiple_assignment_steps() {
    let out = run_prints(
        r#"
program test
    implicit none
    integer :: a
    integer :: b
    a = 1
    b = 2
    a = a + b
    print *, a
    print *, a * b
end program test
"#,
    );
    assert_eq!(out, vec!["3", "6"]);
}

#[test]
fn basics_logical_expression_true_false() {
    let out = run_prints(
        r#"
program test
    implicit none
    integer :: x
    logical :: is_big
    x = 7
    is_big = (x > 5)
    print *, is_big
    print *, .not. is_big
end program test
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}
