! vybe-test: fortran/array_shape_casting_assignments/array_shape_casting_assignments_scalar_to_ranked_array
! origin: languages/fortran/tests/fortran/test_array_shape_casting_assignments.rs

program array_shape_casting_assignments_scalar_to_ranked_array
    integer :: matrix(2, 2)
    matrix = 4
    if ((matrix(1, 1)) /= 4) then
    print *, "FAIL: want [4] got [", matrix(1, 1), "]"
    stop 1
end if
    if ((matrix(2, 2)) /= 4) then
    print *, "FAIL: want [4] got [", matrix(2, 2), "]"
    stop 1
end if
    if ((sum(matrix)) /= 16) then
    print *, "FAIL: want [16] got [", sum(matrix), "]"
    stop 1
end if
end program array_shape_casting_assignments_scalar_to_ranked_array
