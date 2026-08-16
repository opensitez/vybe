! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_non_default_origin_2d
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_non_default_origin_2d
    integer :: matrix(-1:2,-2:1)
    matrix = 1
    matrix(0,0) = 5
    matrix(1:2, -2:-1) = 3
    if ((sum(matrix)) /= 28) then
    print *, "FAIL: want [28] got [", sum(matrix), "]"
    stop 1
end if
    if ((matrix(0,0)) /= 5) then
    print *, "FAIL: want [5] got [", matrix(0,0), "]"
    stop 1
end if
    if ((lbound(matrix, 1)) /= -1) then
    print *, "FAIL: want [-1] got [", lbound(matrix, 1), "]"
    stop 1
end if
    if ((lbound(matrix, 2)) /= -2) then
    print *, "FAIL: want [-2] got [", lbound(matrix, 2), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_non_default_origin_2d
