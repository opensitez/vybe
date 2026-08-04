! vybe-test: fortran/array_subscript_bounds_zero/array_subscript_bounds_zero_basic_bounds
! origin: languages/fortran/tests/fortran/test_array_subscript_bounds_zero.rs

program array_subscript_bounds_zero_basic_bounds
    integer :: values(0:4)
    values = (/10, 11, 12, 13, 14/)
    if ((lbound(values)) /= 0) then
    print *, "FAIL: want [0] got [", lbound(values), "]"
    stop 1
end if
    if ((ubound(values)) /= 4) then
    print *, "FAIL: want [4] got [", ubound(values), "]"
    stop 1
end if
    if ((values(0)) /= 10) then
    print *, "FAIL: want [10] got [", values(0), "]"
    stop 1
end if
    if ((values(4)) /= 14) then
    print *, "FAIL: want [14] got [", values(4), "]"
    stop 1
end if
end program array_subscript_bounds_zero_basic_bounds
