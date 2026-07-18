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
