! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_inside_while_with_bounds
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_inside_while_with_bounds
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
    integer :: a(3)
    integer :: i
    integer :: total
    a = (/ 1, 2, 3 /)
    total = 0
    i = 1
    do while (i <= 5)
        if (i >= lbound(a, 1) .and. i <= ubound(a, 1)) then
            total = total + a(i)
        end if
        i = i + 1
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((total) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program array_bounds_check_failure_paths_inside_while_with_bounds
