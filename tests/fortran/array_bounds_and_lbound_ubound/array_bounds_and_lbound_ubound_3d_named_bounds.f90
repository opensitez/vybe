! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_3d_named_bounds
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_3d_named_bounds
    integer :: cube(-1:1, 2:4, 5:6)
    if ((lbound(cube, 1)) /= -1) then
    print *, "FAIL: want [-1] got [", lbound(cube, 1), "]"
    stop 1
end if
    if ((ubound(cube, 1)) /= 1) then
    print *, "FAIL: want [1] got [", ubound(cube, 1), "]"
    stop 1
end if
    if ((lbound(cube, 2)) /= 2) then
    print *, "FAIL: want [2] got [", lbound(cube, 2), "]"
    stop 1
end if
    if ((ubound(cube, 2)) /= 4) then
    print *, "FAIL: want [4] got [", ubound(cube, 2), "]"
    stop 1
end if
    if ((lbound(cube, 3)) /= 5) then
    print *, "FAIL: want [5] got [", lbound(cube, 3), "]"
    stop 1
end if
    if ((ubound(cube, 3)) /= 6) then
    print *, "FAIL: want [6] got [", ubound(cube, 3), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_3d_named_bounds
