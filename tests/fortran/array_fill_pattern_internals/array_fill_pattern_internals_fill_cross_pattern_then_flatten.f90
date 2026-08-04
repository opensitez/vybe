! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_cross_pattern_then_flatten
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_cross_pattern_then_flatten
    integer :: matrix(4,4)
    integer :: total
    matrix = 0
    matrix(2,:) = 3
    matrix(1,2) = 3
    matrix(3,2) = 3
    matrix(:,3) = 3
    total = sum(matrix)
    if ((total) /= 30) then
    print *, "FAIL: want [30] got [", total, "]"
    stop 1
end if
    if ((matrix(2,2)) /= 3) then
    print *, "FAIL: want [3] got [", matrix(2,2), "]"
    stop 1
end if
    if ((matrix(1,1)) /= 0) then
    print *, "FAIL: want [0] got [", matrix(1,1), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_cross_pattern_then_flatten
