! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_nested_calls_with_invariant_bounds
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program t
    integer :: values(1:7)
    integer :: found

    values = (/ 2, 4, 6, 8, 10, 12, 14 /)
    call check(values, 1, found)
    if ((found) /= 2) then
    print *, "FAIL: want [2] got [", found, "]"
    stop 1
end if
    call check(values, 10, found)
    if ((found) /= -1) then
    print *, "FAIL: want [-1] got [", found, "]"
    stop 1
end if

contains
    subroutine check(a, idx, out)
        integer, intent(in) :: a(:)
        integer, intent(in) :: idx
        integer, intent(out) :: out
        if (idx < lbound(a, 1) .or. idx > ubound(a, 1)) then
            out = -1
        else
            out = a(idx)
        end if
    end subroutine check
end program t
