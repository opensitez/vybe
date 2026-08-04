! vybe-test: fortran/io_advise_and_error_state/test_io_advise_and_error_state_inquire_opened_after_close
! origin: languages/fortran/tests/fortran/test_io_advise_and_error_state.rs

program test_io_advise_and_error_state
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 0 ]
    integer :: unit
    logical :: opened
    open(newunit=unit, status='scratch')
    close(unit)
    inquire(unit=unit, opened=opened)
    if (opened) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((0) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_io_advise_and_error_state
