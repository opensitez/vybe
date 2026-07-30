use super::helpers::run_prints;

#[test]
fn test_io_internal_file_field_sizing_reports_iolength() {
    let out = run_prints(
        r#"
program test_io_internal_file_field_sizing
    integer :: required
    integer :: code
    integer :: value
    value = 256
    inquire(iolength=required) value
    if (required > 0) code = 1
    print *, code
    print *, required
end program test_io_internal_file_field_sizing
"#,
    );

    assert_eq!(out[0], "1");
}

#[test]
fn test_io_internal_file_field_sizing_real_and_logical() {
    let out = run_prints(
        r#"
program test_io_internal_file_field_sizing
    integer :: required
    logical :: good
    real :: value
    logical :: l
    value = 1.5
    l = .true.
    inquire(iolength=required) value, l
    if (required > 0) then
        good = .true.
    else
        good = .false.
    end if
    print *, merge(1, 0, good)
    print *, required
end program test_io_internal_file_field_sizing
"#,
    );

    assert_eq!(out.len(), 2);
}

#[test]
fn test_io_internal_file_field_sizing_character_record() {
    let out = run_prints(
        r#"
program test_io_internal_file_field_sizing
    integer :: n
    character(len=12) :: name
    name = 'fortran_test'
    inquire(iolength=n) name
    print *, n
end program test_io_internal_file_field_sizing
"#,
    );

    assert_eq!(out, vec!["12"]);
}
