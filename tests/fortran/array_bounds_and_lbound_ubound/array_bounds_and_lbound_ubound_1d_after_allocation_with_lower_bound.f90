! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_1d_after_allocation_with_lower_bound
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_1d_after_allocation_with_lower_bound
    integer, allocatable :: values(:)
    allocate(values(-4:1))
    if ((lbound(values, 1)) /= -4) then
    print *, "FAIL: want [-4] got [", lbound(values, 1), "]"
    stop 1
end if
    if ((ubound(values, 1)) /= 1) then
    print *, "FAIL: want [1] got [", ubound(values, 1), "]"
    stop 1
end if
    deallocate(values)
end program array_bounds_and_lbound_ubound_1d_after_allocation_with_lower_bound
