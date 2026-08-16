! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_2d_columns_with_scalar
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_2d_columns_with_scalar
    integer :: matrix(3,3)
    matrix = 1
    matrix(:,2) = -2
    if ((sum(matrix)) /= 0) then
    print *, "FAIL: want [0] got [", sum(matrix), "]"
    stop 1
end if
    if ((matrix(1,2)) /= -2) then
    print *, "FAIL: want [-2] got [", matrix(1,2), "]"
    stop 1
end if
    if ((matrix(3,2)) /= -2) then
    print *, "FAIL: want [-2] got [", matrix(3,2), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_2d_columns_with_scalar
