! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_section_outside_reject
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_section_outside_reject
    integer :: a(1:8)
    integer :: status
    a = (/ (i, i = 1, 8) /)

    if (9 >= lbound(a, 1) .and. 12 <= ubound(a, 1)) then
        status = a(9)
    else
        status = -1
    end if
    if ((status) /= -1) then
    print *, "FAIL: want [-1] got [", status, "]"
    stop 1
end if

    if (7 >= lbound(a, 1) .and. 9 <= ubound(a, 1)) then
        status = a(7)
    else
        status = -1
    end if
    if ((status) /= 7) then
    print *, "FAIL: want [7] got [", status, "]"
    stop 1
end if
end program array_bounds_check_failure_paths_section_outside_reject
