use super::helpers::run_prints;

#[test]
fn test_input_validation_for_read_validates_integer_string() {
    let out = run_prints(
        r#"
program test_input_validation_for_read
    character(len=8) :: src
    integer :: value
    integer :: status
    src = '42'
    read(src, *, iostat=status) value
    print *, status
    print *, value
end program test_input_validation_for_read
"#,
    );

    assert_eq!(out, vec!["0", "42"]);
}

#[test]
fn test_input_validation_for_read_handles_bad_integer_format() {
    let out = run_prints(
        r#"
program test_input_validation_for_read
    character(len=8) :: src
    integer :: value
    integer :: status
    src = 'abc'
    value = 99
    read(src, *, iostat=status) value
    if (status /= 0) then
        print *, 1
    else
        print *, 0
    end if
end program test_input_validation_for_read
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_input_validation_for_read_parses_real_value() {
    let out = run_prints(
        r#"
program test_input_validation_for_read
    character(len=16) :: src
    real :: value
    integer :: status
    src = '1.25e1'
    read(src, *, iostat=status) value
    print *, status
    print *, nint(value)
end program test_input_validation_for_read
"#,
    );

    assert_eq!(out, vec!["0", "13"]);
}

#[test]
fn test_input_validation_for_read_parses_logical_from_string() {
    let out = run_prints(
        r#"
program test_input_validation_for_read
    character(len=8) :: src
    logical :: value
    integer :: status
    src = '.true.'
    read(src, *, iostat=status) value
    print *, status
    print *, value
end program test_input_validation_for_read
"#,
    );

    assert_eq!(out, vec!["0", ".TRUE."]);
}

#[test]
fn test_input_validation_for_read_array_values_with_iostat() {
    let out = run_prints(
        r#"
program test_input_validation_for_read
    character(len=32) :: src
    integer :: values(3)
    integer :: status
    src = '1 2 3'
    read(src, *, iostat=status) values
    print *, status
    print *, sum(values)
end program test_input_validation_for_read
"#,
    );

    assert_eq!(out, vec!["0", "6"]);
}
