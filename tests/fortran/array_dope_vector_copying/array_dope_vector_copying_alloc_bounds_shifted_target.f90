! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_alloc_bounds_shifted_target
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_alloc_bounds_shifted_target
    integer, allocatable :: source(:)
    integer :: dest(-2:2)
    integer :: i
    source = (/ 1, 2, 3, 4, 5 /)
    dest(-2:2) = source
    if ((lbound(dest, 1)) /= -2) then
    print *, "FAIL: want [-2] got [", lbound(dest, 1), "]"
    stop 1
end if
    if ((ubound(dest, 1)) /= 2) then
    print *, "FAIL: want [2] got [", ubound(dest, 1), "]"
    stop 1
end if
    if ((dest(-2)) /= 1) then
    print *, "FAIL: want [1] got [", dest(-2), "]"
    stop 1
end if
    if ((dest(2)) /= 5) then
    print *, "FAIL: want [5] got [", dest(2), "]"
    stop 1
end if
end program array_dope_vector_copying_alloc_bounds_shifted_target
