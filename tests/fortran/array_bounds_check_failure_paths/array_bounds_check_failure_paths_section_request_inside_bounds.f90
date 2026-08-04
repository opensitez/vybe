! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_section_request_inside_bounds
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_section_request_inside_bounds
    integer :: a(1:10)
    integer :: b
    integer :: c
    a = (/ (i, i = 1, 10) /)

    if (1 >= lbound(a, 1) .and. 5 <= ubound(a, 1)) then
        b = a(1)
        c = a(5)
    else
        b = -1
        c = -1
    end if

    if ((b) /= 1) then
    print *, "FAIL: want [1] got [", b, "]"
    stop 1
end if
    if ((c) /= 5) then
    print *, "FAIL: want [5] got [", c, "]"
    stop 1
end if
end program array_bounds_check_failure_paths_section_request_inside_bounds
