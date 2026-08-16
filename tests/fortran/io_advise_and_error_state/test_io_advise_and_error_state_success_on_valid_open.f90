! vybe-test: fortran/io_advise_and_error_state/test_io_advise_and_error_state_success_on_valid_open
! origin: languages/fortran/tests/fortran/test_io_advise_and_error_state.rs

program test_io_advise_and_error_state
    integer :: unit
    integer :: code
    open(newunit=unit, file='valid_probe.dat', status='replace')
    close(unit)
    open(unit=unit, file='valid_probe.dat', status='old', iostat=code)
    if ((code) /= 5002) then
    print *, "FAIL: want [5002] got [", code, "]"
    stop 1
end if
    close(unit)
end program test_io_advise_and_error_state
