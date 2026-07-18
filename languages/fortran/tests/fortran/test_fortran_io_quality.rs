use super::helpers::run_prints;

#[test]
fn io_quality_print_star_default() {
    let out = run_prints(
        r#"
program io_quality_print_star_default
    print *, 42
end program io_quality_print_star_default
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn io_quality_print_format_integer() {
    let out = run_prints(
        r#"
program io_quality_print_format_integer
    integer :: value
    value = 123
    print '(I0)', value
end program io_quality_print_format_integer
"#,
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn io_quality_internal_character_write() {
    let out = run_prints(
        r#"
program io_quality_internal_character_write
    character(len=24) :: text
    integer :: value
    value = 77
    write (text, '(I0)') value
    print *, trim(text)
end program io_quality_internal_character_write
"#,
    );
    assert_eq!(out, vec!["77"]);
}

#[test]
fn io_quality_internal_character_read() {
    let out = run_prints(
        r#"
program io_quality_internal_character_read
    character(len=16) :: text
    integer :: value
    text = '314'
    read (text, '(I0)') value
    print *, value
end program io_quality_internal_character_read
"#,
    );
    assert_eq!(out, vec!["314"]);
}

#[test]
fn io_quality_multi_value_formatting() {
    let out = run_prints(
        r#"
program io_quality_multi_value_formatting
    integer :: a
    integer :: b
    integer :: c
    a = 1
    b = 2
    c = 3
    print '(I0,1x,I0,1x,I0)', a, b, c
end program io_quality_multi_value_formatting
"#,
    );
    assert_eq!(out, vec!["1 2 3"]);
}

#[test]
fn io_quality_real_fixed_output() {
    let out = run_prints(
        r#"
program io_quality_real_fixed_output
    real :: pi
    pi = 3.14
    print '(F6.2)', pi
end program io_quality_real_fixed_output
"#,
    );
    assert_eq!(out, vec![" 3.14"]);
}

#[test]
fn io_quality_logical_output() {
    let out = run_prints(
        r#"
program io_quality_logical_output
    logical :: enabled
    enabled = .true.
    print *, enabled
end program io_quality_logical_output
"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn io_quality_character_output_trim() {
    let out = run_prints(
        r#"
program io_quality_character_output_trim
    character(len=12) :: word
    word = 'fortran'
    print '(A)', trim(word)
end program io_quality_character_output_trim
"#,
    );
    assert_eq!(out, vec!["fortran"]);
}

#[test]
fn io_quality_string_concatenation_print() {
    let out = run_prints(
        r#"
program io_quality_string_concatenation_print
    character(len=20) :: left
    character(len=20) :: right
    left = 'foo'
    right = 'bar'
    print *, trim(left // right)
end program io_quality_string_concatenation_print
"#,
    );
    assert_eq!(out, vec!["foobar"]);
}

#[test]
fn io_quality_labelled_write_field_width() {
    let out = run_prints(
        r#"
program io_quality_labelled_write_field_width
    integer :: value
    value = 42
    write (*, '(I4)') value
end program io_quality_labelled_write_field_width
"#,
    );
    assert_eq!(out, vec!["  42"]);
}

#[test]
fn io_quality_repeat_format() {
    let out = run_prints(
        r#"
program io_quality_repeat_format
    integer :: i
    character(len=40) :: text
    text = 'ok '
    write (text, '(I0,A)') 4, text
    print *, trim(text)
end program io_quality_repeat_format
"#,
    );
    assert_eq!(out, vec!["4ok "]);
}

#[test]
fn io_quality_parentheses_expression_print() {
    let out = run_prints(
        r#"
program io_quality_parentheses_expression_print
    print *, 2 * (3 + 1)
end program io_quality_parentheses_expression_print
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn io_quality_float_scientific_format() {
    let out = run_prints(
        r#"
program io_quality_float_scientific_format
    real :: value
    value = 1.25
    print '(E10.3)', value
end program io_quality_float_scientific_format
"#,
    );
    assert_eq!(out, vec!["0.125E+01"]);
}
