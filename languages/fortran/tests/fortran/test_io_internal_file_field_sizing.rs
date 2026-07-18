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
