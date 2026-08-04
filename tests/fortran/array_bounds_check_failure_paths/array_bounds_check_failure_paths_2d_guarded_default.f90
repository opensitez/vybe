! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_2d_guarded_default
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_2d_guarded_default
    integer :: matrix(4,4)
    integer :: i, j, hits

    matrix = reshape((/ (i, i = 1, 16) /), (/4,4/))
    hits = 0
    i = -1
    j = 2
    if (i >= 1 .and. i <= 4 .and. j >= 1 .and. j <= 4) hits = hits + 1
    if (i == 1 .and. j == 2) hits = hits + 1
    i = 2
    if (i >= 1 .and. i <= 4 .and. j >= 1 .and. j <= 4) hits = hits + 1
    if ((hits) /= 2) then
    print *, "FAIL: want [2] got [", hits, "]"
    stop 1
end if
end program array_bounds_check_failure_paths_2d_guarded_default
