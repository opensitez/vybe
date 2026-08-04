! vybe-test: fortran/array_subscript_bounds_zero/array_subscript_bounds_zero_matrix_zero_row
! origin: languages/fortran/tests/fortran/test_array_subscript_bounds_zero.rs

program array_subscript_bounds_zero_matrix_zero_row
    integer :: matrix(0:2, 0:2)
    integer :: r
    matrix = reshape((/1, 2, 3, 4, 5, 6, 7, 8, 9/), (/3, 3/))
    r = matrix(0, 0) + matrix(2, 2)
    if ((r) /= 10) then
    print *, "FAIL: want [10] got [", r, "]"
    stop 1
end if
    if ((lbound(matrix, 1)) /= 0) then
    print *, "FAIL: want [0] got [", lbound(matrix, 1), "]"
    stop 1
end if
    if ((lbound(matrix, 2)) /= 0) then
    print *, "FAIL: want [0] got [", lbound(matrix, 2), "]"
    stop 1
end if
end program array_subscript_bounds_zero_matrix_zero_row
