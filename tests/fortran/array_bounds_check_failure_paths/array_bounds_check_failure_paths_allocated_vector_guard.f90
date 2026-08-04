! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_allocated_vector_guard
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_allocated_vector_guard
    integer, allocatable :: a(:)
    integer :: value

    allocate(a(2:6))
    a = (/ 4, 8, 15, 16, 23 /)
    if (lbound(a, 1) /= 2 .or. ubound(a, 1) /= 6) then
        value = -1
    else
        value = a(4)
    end if
    if ((value) /= 15) then
    print *, "FAIL: want [15] got [", value, "]"
    stop 1
end if
    deallocate(a)
end program array_bounds_check_failure_paths_allocated_vector_guard
