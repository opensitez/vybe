! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_from_reshape_then_assign_back
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_from_reshape_then_assign_back
    integer :: matrix(3,2)
    integer, allocatable :: flat(:)
    matrix = reshape((/ 9, 8, 7, 6, 5, 4 /), (/3,2/))
    flat = reshape(matrix, (/6/))
    if ((size(flat)) /= 6) then
    print *, "FAIL: want [6] got [", size(flat), "]"
    stop 1
end if
    if ((sum(flat)) /= 39) then
    print *, "FAIL: want [39] got [", sum(flat), "]"
    stop 1
end if
    if ((flat(1)) /= 9) then
    print *, "FAIL: want [9] got [", flat(1), "]"
    stop 1
end if
    if ((flat(6)) /= 4) then
    print *, "FAIL: want [4] got [", flat(6), "]"
    stop 1
end if
end program array_dope_vector_copying_copy_from_reshape_then_assign_back
