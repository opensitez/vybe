! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_slice_update_with_expression_copy
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_slice_update_with_expression_copy
    integer :: source(1:5)
    integer :: target(1:5)
    source = (/ 2, 4, 6, 8, 10 /)
    target = source
    target(2:4) = target(1:3) + 1
    if ((target(1)) /= 2) then
    print *, "FAIL: want [2] got [", target(1), "]"
    stop 1
end if
    if ((target(2)) /= 3) then
    print *, "FAIL: want [3] got [", target(2), "]"
    stop 1
end if
    if ((target(3)) /= 5) then
    print *, "FAIL: want [5] got [", target(3), "]"
    stop 1
end if
    if ((target(4)) /= 7) then
    print *, "FAIL: want [7] got [", target(4), "]"
    stop 1
end if
    if ((target(5)) /= 10) then
    print *, "FAIL: want [10] got [", target(5), "]"
    stop 1
end if
end program array_dope_vector_copying_slice_update_with_expression_copy
