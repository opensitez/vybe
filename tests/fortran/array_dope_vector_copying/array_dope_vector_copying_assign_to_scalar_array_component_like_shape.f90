! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_assign_to_scalar_array_component_like_shape
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program t
    integer, allocatable :: source(:)
    integer :: target(0:3)
    source = (/ 1, 2, 3, 4 /)
    target = source
    if ((lbound(target, 1)) /= 0) then
    print *, "FAIL: want [0] got [", lbound(target, 1), "]"
    stop 1
end if
    if ((ubound(target, 1)) /= 3) then
    print *, "FAIL: want [3] got [", ubound(target, 1), "]"
    stop 1
end if
    if ((target(0)) /= 1) then
    print *, "FAIL: want [1] got [", target(0), "]"
    stop 1
end if
    if ((target(3)) /= 4) then
    print *, "FAIL: want [4] got [", target(3), "]"
    stop 1
end if
end program t
