! vybe-test: fortran/array_shape_casting_assignments/array_shape_casting_assignments_matrix_to_vector_via_reshape
! origin: languages/fortran/tests/fortran/test_array_shape_casting_assignments.rs

program array_shape_casting_assignments_matrix_to_vector_via_reshape
    integer :: matrix(3, 2)
    integer :: flat(6)
    matrix = reshape((/1, 2, 3, 4, 5, 6/), (/3, 2/))
    flat = reshape(matrix, (/6/))
    if ((flat(1)) /= 1) then
    print *, "FAIL: want [1] got [", flat(1), "]"
    stop 1
end if
    if ((flat(6)) /= 6) then
    print *, "FAIL: want [6] got [", flat(6), "]"
    stop 1
end if
    if ((sum(flat)) /= 21) then
    print *, "FAIL: want [21] got [", sum(flat), "]"
    stop 1
end if
end program array_shape_casting_assignments_matrix_to_vector_via_reshape
