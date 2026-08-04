! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_after_realloc_grows
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_after_realloc_grows
    integer, allocatable :: values(:)
    values = (/ 1, 2, 3 /)
    values = 0
    values = (/ 4, 4 /)
    values = 3
    if ((size(values)) /= 2) then
    print *, "FAIL: want [2] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 6) then
    print *, "FAIL: want [6] got [", sum(values), "]"
    stop 1
end if
    if ((values(2)) /= 3) then
    print *, "FAIL: want [3] got [", values(2), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_after_realloc_grows
