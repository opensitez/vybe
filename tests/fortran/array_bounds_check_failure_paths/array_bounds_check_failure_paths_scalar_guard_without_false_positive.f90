! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_scalar_guard_without_false_positive
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 165 ]
    integer :: a(3:7)
    integer :: idx
    integer :: count

    a = (/ 11, 22, 33, 44, 55 /)
    count = 0
    do idx = 3, 7
        if (idx >= lbound(a, 1) .and. idx <= ubound(a, 1)) then
            count = count + a(idx)
        end if
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((count) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", count, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
