use super::helpers::run_prints;

#[test]
fn test_file_open_status_modes_scratch_unit_opens_and_closes() {
    let out = run_prints(
        r#"
program test_file_open_status_modes
    integer :: unit
    integer :: code
    open(newunit=unit, status='scratch', action='readwrite', iostat=code)
    if (code == 0) then
        print *, 1
    else
        print *, 0
    end if
    close(unit)
    print *, code
end program test_file_open_status_modes
"#,
    );

    assert_eq!(out, vec!["1", "0"]);
}
