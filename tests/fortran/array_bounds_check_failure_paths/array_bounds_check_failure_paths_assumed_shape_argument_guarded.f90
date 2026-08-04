! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_assumed_shape_argument_guarded
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_assumed_shape_argument_guarded
    integer :: source(1:4)
    integer :: value
    source = (/ 4, 3, 2, 1 /)
    call read_with_guard(source, 1, value)
    if ((value) /= 4) then
    print *, "FAIL: want [4] got [", value, "]"
    stop 1
end if
    call read_with_guard(source, 8, value)
    if ((value) /= -1) then
    print *, "FAIL: want [-1] got [", value, "]"
    stop 1
end if

contains
    subroutine read_with_guard(items, idx, value)
        integer, intent(in) :: items(:)
        integer, intent(in) :: idx
        integer, intent(out) :: value
        if (idx < lbound(items, 1) .or. idx > ubound(items, 1)) then
            value = -1
        else
            value = items(idx)
        end if
    end subroutine read_with_guard
end program array_bounds_check_failure_paths_assumed_shape_argument_guarded
