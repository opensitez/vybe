! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_slice_pointer_no_dim
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_slice_pointer_no_dim
    integer, target :: source(-4:4)
    integer, pointer :: alias(:)
    integer :: lb(1), ub(1)
    alias => source(2:4)
    lb = lbound(alias)
    ub = ubound(alias)
    if ((lb(1)) /= 2) then
    print *, "FAIL: want [2] got [", lb(1), "]"
    stop 1
end if
    if ((ub(1)) /= 4) then
    print *, "FAIL: want [4] got [", ub(1), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_slice_pointer_no_dim
