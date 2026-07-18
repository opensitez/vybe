use super::helpers::run_prints;

#[test]
fn test_io_advise_and_error_state_sets_iostat() {
    let out = run_prints(
        r#"
program test_io_advise_and_error_state
    integer :: unit
    integer :: code
    open(newunit=unit, status='scratch', action='readwrite')
    close(unit)
    open(unit=unit, file='nope', status='old', iostat=code)
    print *, code
end program test_io_advise_and_error_state
"#,
    );

    assert_eq!(out.len(), 1);
}
