! vybe-test: fortran/coarray_sync_errors/test_coarray_sync_errors_checks_non_terminating_errors
! origin: languages/fortran/tests/fortran/test_coarray_sync_errors.rs

program test_coarray_sync_errors
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 0 ]
    integer :: status
    status = 0
    if (status == 0) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((0) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_coarray_sync_errors
