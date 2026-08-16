! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_matrix_outer_border
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_matrix_outer_border
    integer :: matrix(3,4)
    integer :: perimeter
    matrix = 0
    matrix(1,:) = 9
    matrix(3,:) = 9
    matrix(2,1) = 9
    matrix(2,4) = 9
    perimeter = sum(matrix)
    if ((perimeter) /= 90) then
    print *, "FAIL: want [90] got [", perimeter, "]"
    stop 1
end if
    if ((matrix(2,2)) /= 0) then
    print *, "FAIL: want [0] got [", matrix(2,2), "]"
    stop 1
end if
    if ((matrix(1,2)) /= 9) then
    print *, "FAIL: want [9] got [", matrix(1,2), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_matrix_outer_border
