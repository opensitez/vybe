! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_slice_by_variable_boundaries
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_slice_by_variable_boundaries
    integer :: values(1:10)
    integer :: start_idx
    integer :: end_idx
    values = 2
    start_idx = 4
    end_idx = 8
    values(start_idx:end_idx) = -1
    if ((sum(values)) /= 5) then
    print *, "FAIL: want [5] got [", sum(values), "]"
    stop 1
end if
    if ((values(start_idx)) /= -1) then
    print *, "FAIL: want [-1] got [", values(start_idx), "]"
    stop 1
end if
    if ((values(end_idx)) /= -1) then
    print *, "FAIL: want [-1] got [", values(end_idx), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_slice_by_variable_boundaries
