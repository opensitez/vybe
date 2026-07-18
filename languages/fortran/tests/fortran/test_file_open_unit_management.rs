use super::helpers::run_prints;

#[test]
fn test_file_open_unit_management_tracks_reopen_cycle() {
    let out = run_prints(
        r#"
program test_file_open_unit_management
    integer :: first_unit
    integer :: second_unit
    integer :: status
    open(newunit=first_unit, status='scratch', action='readwrite')
    close(first_unit)
    open(newunit=second_unit, status='scratch', action='readwrite', iostat=status)
    print *, first_unit
    print *, status
    close(second_unit)
end program test_file_open_unit_management
"#,
    );

    assert_eq!(out.len(), 2);
    assert_eq!(out[0], out[1]);
}
