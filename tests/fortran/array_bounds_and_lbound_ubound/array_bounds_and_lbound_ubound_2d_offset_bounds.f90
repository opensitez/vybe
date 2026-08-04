! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_2d_offset_bounds
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_2d_offset_bounds
    integer :: grid(-2:2, 10:12)
    if ((lbound(grid, 1)) /= -2) then
    print *, "FAIL: want [-2] got [", lbound(grid, 1), "]"
    stop 1
end if
    if ((ubound(grid, 1)) /= 2) then
    print *, "FAIL: want [2] got [", ubound(grid, 1), "]"
    stop 1
end if
    if ((lbound(grid, 2)) /= 10) then
    print *, "FAIL: want [10] got [", lbound(grid, 2), "]"
    stop 1
end if
    if ((ubound(grid, 2)) /= 12) then
    print *, "FAIL: want [12] got [", ubound(grid, 2), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_2d_offset_bounds
