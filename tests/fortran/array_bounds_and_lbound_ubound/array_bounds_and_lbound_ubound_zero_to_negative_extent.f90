! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_zero_to_negative_extent
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_zero_to_negative_extent
    integer :: values(-2:-2)
    if ((lbound(values, 1)) /= -2) then
    print *, "FAIL: want [-2] got [", lbound(values, 1), "]"
    stop 1
end if
    if ((ubound(values, 1)) /= -2) then
    print *, "FAIL: want [-2] got [", ubound(values, 1), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_zero_to_negative_extent
