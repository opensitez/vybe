! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_with_explicit_shape_declaration
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_with_explicit_shape_declaration
    integer :: source(1:3)
    integer :: target(4:6)
    source = (/ 3, 6, 9 /)
    target = source
    if ((lbound(source, 1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(source, 1), "]"
    stop 1
end if
    if ((ubound(target, 1)) /= 6) then
    print *, "FAIL: want [6] got [", ubound(target, 1), "]"
    stop 1
end if
    if ((target(4)) /= 3) then
    print *, "FAIL: want [3] got [", target(4), "]"
    stop 1
end if
    if ((target(6)) /= 9) then
    print *, "FAIL: want [9] got [", target(6), "]"
    stop 1
end if
end program array_dope_vector_copying_copy_with_explicit_shape_declaration
