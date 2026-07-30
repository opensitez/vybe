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

#[test]
fn test_io_advise_and_error_state_success_on_valid_open() {
    let out = run_prints(
        r#"
program test_io_advise_and_error_state
    integer :: unit
    integer :: code
    open(newunit=unit, file='valid_probe.dat', status='replace')
    close(unit)
    open(unit=unit, file='valid_probe.dat', status='old', iostat=code)
    print *, code
    close(unit)
end program test_io_advise_and_error_state
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_io_advise_and_error_state_endfile_sets_end() {
    let out = run_prints(
        r#"
program test_io_advise_and_error_state
    integer :: unit
    integer :: n
    integer :: ios
    character(len=12) :: buf
    open(newunit=unit, file='end_probe.dat', status='replace')
    write(unit, '(I0)') 7
    rewind(unit)
    read(unit, *, iostat=ios) n
    read(unit, *, iostat=ios) n
    if (ios < 0) then
        print *, 1
    else
        print *, 0
    end if
    close(unit, status='delete')
end program test_io_advise_and_error_state
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_io_advise_and_error_state_inquire_opened_after_close() {
    let out = run_prints(
        r#"
program test_io_advise_and_error_state
    integer :: unit
    logical :: opened
    open(newunit=unit, status='scratch')
    close(unit)
    inquire(unit=unit, opened=opened)
    if (opened) then
        print *, 1
    else
        print *, 0
    end if
end program test_io_advise_and_error_state
"#,
    );

    assert_eq!(out, vec!["0"]);
}
