! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_1d_default_bounds
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_1d_default_bounds
    integer :: values(6)
    if ((lbound(values, 1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(values, 1), "]"
    stop 1
end if
    if ((ubound(values, 1)) /= 6) then
    print *, "FAIL: want [6] got [", ubound(values, 1), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_1d_default_bounds
