! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_from_constructor_then_merge
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_from_constructor_then_merge
    integer, allocatable :: values(:), mask(:)
    values = (/ 1, 2, 3, 4, 5 /)
    mask = merge(1, 0, values > 2)
    values = merge(100, values, mask == 1)
    if ((sum(values)) /= 315) then
    print *, "FAIL: want [315] got [", sum(values), "]"
    stop 1
end if
    if ((values(2)) /= 2) then
    print *, "FAIL: want [2] got [", values(2), "]"
    stop 1
end if
    if ((values(5)) /= 100) then
    print *, "FAIL: want [100] got [", values(5), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_from_constructor_then_merge
