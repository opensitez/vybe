! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_zero_size_array
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_zero_size_array
    integer :: values(0:0)
    if ((lbound(values, 1)) /= 0) then
    print *, "FAIL: want [0] got [", lbound(values, 1), "]"
    stop 1
end if
    if ((ubound(values, 1)) /= 0) then
    print *, "FAIL: want [0] got [", ubound(values, 1), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_zero_size_array
