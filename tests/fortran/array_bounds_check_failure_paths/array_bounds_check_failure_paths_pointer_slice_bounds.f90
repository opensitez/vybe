! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_pointer_slice_bounds
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_pointer_slice_bounds
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 2, 8 ]
    integer, target :: source(0:9)
    integer, pointer :: alias(:)

    source = (/ (i, i = 0, 9) /)
    alias => source(2:8)

    if (lbound(alias, 1) >= 2 .and. ubound(alias, 1) <= 8) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((alias(lbound(alias, 1))) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", alias(lbound(alias, 1)), "]"
            stop 1
        end if
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((alias(ubound(alias, 1))) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", alias(ubound(alias, 1)), "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((-1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", -1, "]"
            stop 1
        end if
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((-1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", -1, "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program array_bounds_check_failure_paths_pointer_slice_bounds
