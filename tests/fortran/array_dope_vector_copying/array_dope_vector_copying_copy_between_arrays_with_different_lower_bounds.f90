! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_between_arrays_with_different_lower_bounds
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program t
    integer :: left(-1:4)
    integer :: right(10:15)
    left = (/ 1, 2, 3, 4, 5, 6 /)
    right = left
    if ((lbound(left, 1)) /= -1) then
    print *, "FAIL: want [-1] got [", lbound(left, 1), "]"
    stop 1
end if
    if ((lbound(right, 1)) /= 10) then
    print *, "FAIL: want [10] got [", lbound(right, 1), "]"
    stop 1
end if
    if ((right(10)) /= 1) then
    print *, "FAIL: want [1] got [", right(10), "]"
    stop 1
end if
    if ((right(15)) /= 6) then
    print *, "FAIL: want [6] got [", right(15), "]"
    stop 1
end if
end program t
