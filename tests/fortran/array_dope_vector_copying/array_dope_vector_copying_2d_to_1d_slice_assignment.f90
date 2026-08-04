! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_2d_to_1d_slice_assignment
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_2d_to_1d_slice_assignment
    integer :: matrix(2,3)
    integer :: column(2)
    matrix = reshape((/ 1, 2, 3, 4, 5, 6 /), (/2,3/))
    column = matrix(:,2)
    if ((size(column)) /= 2) then
    print *, "FAIL: want [2] got [", size(column), "]"
    stop 1
end if
    if ((sum(column)) /= 7) then
    print *, "FAIL: want [7] got [", sum(column), "]"
    stop 1
end if
    if ((column(1)) /= 2) then
    print *, "FAIL: want [2] got [", column(1), "]"
    stop 1
end if
    if ((column(2)) /= 4) then
    print *, "FAIL: want [4] got [", column(2), "]"
    stop 1
end if
end program array_dope_vector_copying_2d_to_1d_slice_assignment
