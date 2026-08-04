! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_allocatable_slice_after_assign
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_allocatable_slice_after_assign
    integer, allocatable :: buffer(:)
    integer, allocatable :: slice(:)
    allocate(buffer(1:12))
    slice => buffer(4:9)
    if ((lbound(slice, 1)) /= 4) then
    print *, "FAIL: want [4] got [", lbound(slice, 1), "]"
    stop 1
end if
    if ((ubound(slice, 1)) /= 9) then
    print *, "FAIL: want [9] got [", ubound(slice, 1), "]"
    stop 1
end if
    deallocate(buffer)
end program array_bounds_and_lbound_ubound_allocatable_slice_after_assign
