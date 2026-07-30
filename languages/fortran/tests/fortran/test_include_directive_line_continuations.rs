use super::helpers::run_prints;

#[test]
fn test_include_directive_line_continuations_use_ampersand() {
    let out = run_prints(
        r#"
program test_include_directive_line_continuations
    integer :: value
    value = 1 + 2 + &
            3 + 4
    print *, value
end program test_include_directive_line_continuations
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_include_directive_line_continuations_char_concat_split_across_lines() {
    let out = run_prints(
        r#"
program test_include_directive_line_continuations
    character(len=20) :: word
    word = "for" // &
           "tran"
    print *, word
end program test_include_directive_line_continuations
"#,
    );

    assert_eq!(out, vec!["fortran"]);
}

#[test]
fn test_include_directive_line_continuations_in_do_header() {
    let out = run_prints(
        r#"
program test_include_directive_line_continuations
    integer :: s
    s = 0
    do i = 1, 4, &
       1
        s = s + i
    end do
    print *, s
end program test_include_directive_line_continuations
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_include_directive_line_continuations_nested_expression_chain() {
    let out = run_prints(
        r#"
program test_include_directive_line_continuations
    integer :: a, b, c
    a = 1 + &
        2 + 3 + &
        4
    b = (a * 2) / &
        (1 + 3)
    c = a + b - 2
    print *, c
end program test_include_directive_line_continuations
"#,
    );

    assert_eq!(out, vec!["13"]);
}

#[test]
fn test_include_directive_line_continuations_in_array_constructor() {
    let out = run_prints(
        r#"
program test_include_directive_line_continuations
    integer :: a(3)
    a = [ 1,  &
          2,  &
          3 ]
    print *, sum(a)
end program test_include_directive_line_continuations
"#,
    );

    assert_eq!(out, vec!["6"]);
}
