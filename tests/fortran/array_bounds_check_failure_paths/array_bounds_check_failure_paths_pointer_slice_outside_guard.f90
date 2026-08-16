! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_pointer_slice_outside_guard
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_pointer_slice_outside_guard
    integer, target :: source(10:14)
    integer, pointer :: alias(:)
    integer :: status

    source = (/ 1, 2, 3, 4, 5 /)
    alias => source(10:14)

    if (11 >= lbound(alias, 1) .and. 11 <= ubound(alias, 1)) then
        status = alias(11)
    else
        status = -1
    end if
    if ((status) /= -1) then
    print *, "FAIL: want [-1] got [", status, "]"
    stop 1
end if

    if (15 >= lbound(alias, 1) .and. 15 <= ubound(alias, 1)) then
        status = alias(15)
    else
        status = -2
    end if
    if ((status) /= -2) then
    print *, "FAIL: want [-2] got [", status, "]"
    stop 1
end if
end program array_bounds_check_failure_paths_pointer_slice_outside_guard
