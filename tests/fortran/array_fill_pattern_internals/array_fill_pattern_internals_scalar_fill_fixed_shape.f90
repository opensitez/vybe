! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_scalar_fill_fixed_shape
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_scalar_fill_fixed_shape
    integer :: values(4)
    values = 5
    if ((sum(values)) /= 20) then
    print *, "FAIL: want [20] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 5) then
    print *, "FAIL: want [5] got [", values(1), "]"
    stop 1
end if
    if ((values(4)) /= 5) then
    print *, "FAIL: want [5] got [", values(4), "]"
    stop 1
end if
end program array_fill_pattern_internals_scalar_fill_fixed_shape
