! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_2d_default_bounds
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_2d_default_bounds
    integer :: grid(3, 4)
    if ((lbound(grid, 1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(grid, 1), "]"
    stop 1
end if
    if ((ubound(grid, 1)) /= 3) then
    print *, "FAIL: want [3] got [", ubound(grid, 1), "]"
    stop 1
end if
    if ((lbound(grid, 2)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(grid, 2), "]"
    stop 1
end if
    if ((ubound(grid, 2)) /= 4) then
    print *, "FAIL: want [4] got [", ubound(grid, 2), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_2d_default_bounds
