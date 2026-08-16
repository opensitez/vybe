! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_pointer_slice
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_pointer_slice
    integer, target :: source(-2:6)
    integer, pointer :: alias(:)
    alias => source(-2:6)
    if ((lbound(alias, 1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(alias, 1), "]"
    stop 1
end if
    if ((ubound(alias, 1)) /= 9) then
    print *, "FAIL: want [9] got [", ubound(alias, 1), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_pointer_slice
