! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_multi_dim_guarded_access
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_multi_dim_guarded_access
    integer :: grid(2,3)
    integer :: status
    integer :: i
    integer :: j
    integer :: sum

    grid = reshape((/ 1, 2, 3, 4, 5, 6 /), (/2, 3/))
    sum = 0
    i = 3
    j = 1
    if (i >= lbound(grid, 1) .and. i <= ubound(grid, 1) &
        .and. j >= lbound(grid, 2) .and. j <= ubound(grid, 2)) then
        status = grid(i, j)
    else
        status = -1
    end if
    sum = sum + status

    i = 1
    j = 5
    if (i >= lbound(grid, 1) .and. i <= ubound(grid, 1) &
        .and. j >= lbound(grid, 2) .and. j <= ubound(grid, 2)) then
        status = grid(i, j)
    else
        status = -2
    end if
    sum = sum + status

    if ((sum) /= -3) then
    print *, "FAIL: want [-3] got [", sum, "]"
    stop 1
end if
end program array_bounds_check_failure_paths_multi_dim_guarded_access
