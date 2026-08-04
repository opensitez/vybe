! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_section_start_greater_than_end
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_section_start_greater_than_end
    integer :: a(1:6)
    integer :: sum

    a = (/ 1, 2, 3, 4, 5, 6 /)
    sum = 0
    if (5 >= lbound(a, 1) .and. 3 <= ubound(a, 1) .and. 5 <= 3) then
        sum = sum + a(5) + a(3)
    else
        sum = -99
    end if
    if ((sum) /= -99) then
    print *, "FAIL: want [-99] got [", sum, "]"
    stop 1
end if
end program array_bounds_check_failure_paths_section_start_greater_than_end
