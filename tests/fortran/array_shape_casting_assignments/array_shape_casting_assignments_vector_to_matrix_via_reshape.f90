! vybe-test: fortran/array_shape_casting_assignments/array_shape_casting_assignments_vector_to_matrix_via_reshape
! origin: languages/fortran/tests/fortran/test_array_shape_casting_assignments.rs

program array_shape_casting_assignments_vector_to_matrix_via_reshape
    integer :: flat(6)
    integer :: matrix(2, 3)
    flat = (/1, 2, 3, 4, 5, 6/)
    matrix = reshape(flat, (/2, 3/))
    if ((matrix(1, 1)) /= 1) then
    print *, "FAIL: want [1] got [", matrix(1, 1), "]"
    stop 1
end if
    if ((matrix(2, 3)) /= 6) then
    print *, "FAIL: want [6] got [", matrix(2, 3), "]"
    stop 1
end if
    if ((sum(matrix)) /= 21) then
    print *, "FAIL: want [21] got [", sum(matrix), "]"
    stop 1
end if
end program array_shape_casting_assignments_vector_to_matrix_via_reshape
