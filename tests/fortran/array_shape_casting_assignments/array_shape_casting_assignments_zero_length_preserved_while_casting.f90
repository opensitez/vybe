! vybe-test: fortran/array_shape_casting_assignments/array_shape_casting_assignments_zero_length_preserved_while_casting
! origin: languages/fortran/tests/fortran/test_array_shape_casting_assignments.rs

program t
    integer :: flat(0)
    integer :: matrix(0, 1)
    matrix = reshape(flat, (/0, 1/))
    if ((size(matrix, 1)) /= 0) then
    print *, "FAIL: want [0] got [", size(matrix, 1), "]"
    stop 1
end if
    if ((size(matrix, 2)) /= 1) then
    print *, "FAIL: want [1] got [", size(matrix, 2), "]"
    stop 1
end if
end program t
