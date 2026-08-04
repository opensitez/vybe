! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_vectorized_fill_via_repeat_expression
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_vectorized_fill_via_repeat_expression
    integer, allocatable :: values(:)
    values = (/ 3 * 0 /)
    if ((size(values)) /= 3) then
    print *, "FAIL: want [3] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 0) then
    print *, "FAIL: want [0] got [", sum(values), "]"
    stop 1
end if
    values = 11
    if ((sum(values)) /= 33) then
    print *, "FAIL: want [33] got [", sum(values), "]"
    stop 1
end if
    if ((values(3)) /= 11) then
    print *, "FAIL: want [11] got [", values(3), "]"
    stop 1
end if
end program array_fill_pattern_internals_vectorized_fill_via_repeat_expression
