! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_2d_rows_with_scalar
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_2d_rows_with_scalar
    integer :: matrix(3,3)
    matrix = 0
    matrix(2,:) = 4
    if ((sum(matrix)) /= 12) then
    print *, "FAIL: want [12] got [", sum(matrix), "]"
    stop 1
end if
    if ((matrix(2,1)) /= 4) then
    print *, "FAIL: want [4] got [", matrix(2,1), "]"
    stop 1
end if
    if ((matrix(1,1)) /= 0) then
    print *, "FAIL: want [0] got [", matrix(1,1), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_2d_rows_with_scalar
