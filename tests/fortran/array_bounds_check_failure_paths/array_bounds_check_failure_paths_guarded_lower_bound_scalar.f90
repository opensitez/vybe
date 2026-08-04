! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_guarded_lower_bound_scalar
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_guarded_lower_bound_scalar
    integer :: a(5)
    integer :: index
    integer :: status
    a = (/ 10, 20, 30, 40, 50 /)
    index = 0
    if (index < 1) then
        status = 0
    else
        status = a(index)
    end if
    if ((status) /= 0) then
    print *, "FAIL: want [0] got [", status, "]"
    stop 1
end if
end program array_bounds_check_failure_paths_guarded_lower_bound_scalar
