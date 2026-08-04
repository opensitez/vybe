! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_scalar_fill_allocation
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_scalar_fill_allocation
    integer, allocatable :: values(:)
    allocate(values(1:6))
    values = 7
    if ((sum(values)) /= 42) then
    print *, "FAIL: want [42] got [", sum(values), "]"
    stop 1
end if
    if ((minval(values)) /= 7) then
    print *, "FAIL: want [7] got [", minval(values), "]"
    stop 1
end if
    if ((maxval(values)) /= 7) then
    print *, "FAIL: want [7] got [", maxval(values), "]"
    stop 1
end if
end program array_fill_pattern_internals_scalar_fill_allocation
