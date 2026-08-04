! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_allocated_vector_misaligned_request
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_allocated_vector_misaligned_request
    integer, allocatable :: a(:)
    integer :: result
    allocate(a(-1:1))
    a = (/ 7, 8, 9 /)

    if (a(0) == 8 .and. lbound(a, 1) == -1) then
        result = 1
    else
        result = 0
    end if
    if ((result) /= 1) then
    print *, "FAIL: want [1] got [", result, "]"
    stop 1
end if

    if (ubound(a, 1) == 1 .and. lbound(a, 1) == -1) then
        result = result + 1
    else
        result = result - 1
    end if
    if ((result) /= 2) then
    print *, "FAIL: want [2] got [", result, "]"
    stop 1
end if

    deallocate(a)
end program array_bounds_check_failure_paths_allocated_vector_misaligned_request
