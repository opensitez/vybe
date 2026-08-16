! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_referenced_from_nested_array
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_referenced_from_nested_array
    integer :: values(-3:3)
    integer :: nested(2)
    nested(1) = -3
    nested(2) = 3
    if ((lbound(values(nested(1):nested(2)), 1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(values(nested(1):nested(2)), 1), "]"
    stop 1
end if
    if ((ubound(values(nested(1):nested(2)), 1)) /= 7) then
    print *, "FAIL: want [7] got [", ubound(values(nested(1):nested(2)), 1), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_referenced_from_nested_array
