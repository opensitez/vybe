! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_slice_by_constant_boundaries
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_slice_by_constant_boundaries
    integer :: values(1:8)
    values = 1
    values(3:6) = 9
    if ((sum(values)) /= 40) then
    print *, "FAIL: want [40] got [", sum(values), "]"
    stop 1
end if
    if ((lbound(values(3:6), 1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(values(3:6), 1), "]"
    stop 1
end if
    if ((ubound(values(3:6), 1)) /= 4) then
    print *, "FAIL: want [4] got [", ubound(values(3:6), 1), "]"
    stop 1
end if
    if ((values(5)) /= 9) then
    print *, "FAIL: want [9] got [", values(5), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_slice_by_constant_boundaries
