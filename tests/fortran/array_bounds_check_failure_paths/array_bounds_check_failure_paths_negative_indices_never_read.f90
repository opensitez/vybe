! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_negative_indices_never_read
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_negative_indices_never_read
    integer :: values(-1:2)
    integer :: result
    integer :: idx
    values(-1) = 9
    values(0) = 8
    values(1) = 7
    values(2) = 6

    result = 0
    idx = -2
    if (idx >= lbound(values, 1) .and. idx <= ubound(values, 1)) then
        result = values(idx)
    else
        result = -1
    end if
    if ((result) /= -1) then
    print *, "FAIL: want [-1] got [", result, "]"
    stop 1
end if

    idx = 3
    if (idx >= lbound(values, 1) .and. idx <= ubound(values, 1)) then
        result = values(idx)
    else
        result = -2
    end if
    if ((result) /= -2) then
    print *, "FAIL: want [-2] got [", result, "]"
    stop 1
end if
end program array_bounds_check_failure_paths_negative_indices_never_read
