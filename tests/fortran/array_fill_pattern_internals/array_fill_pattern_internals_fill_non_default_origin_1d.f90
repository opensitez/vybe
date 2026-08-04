! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_non_default_origin_1d
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_non_default_origin_1d
    integer :: values(-2:2)
    values = 0
    values(-2:0) = 4
    values(1:2) = -1
    if ((sum(values)) /= 11) then
    print *, "FAIL: want [11] got [", sum(values), "]"
    stop 1
end if
    if ((lbound(values,1)) /= -2) then
    print *, "FAIL: want [-2] got [", lbound(values,1), "]"
    stop 1
end if
    if ((ubound(values,1)) /= 2) then
    print *, "FAIL: want [2] got [", ubound(values,1), "]"
    stop 1
end if
    if ((values(0)) /= 4) then
    print *, "FAIL: want [4] got [", values(0), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_non_default_origin_1d
