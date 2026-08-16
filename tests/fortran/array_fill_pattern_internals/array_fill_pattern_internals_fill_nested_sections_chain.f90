! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_nested_sections_chain
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_nested_sections_chain
    integer :: matrix(4,4)
    matrix = 1
    matrix(2:3,2:3) = 7
    matrix(3,1) = 9
    if ((sum(matrix)) /= 48) then
    print *, "FAIL: want [48] got [", sum(matrix), "]"
    stop 1
end if
    if ((matrix(2,2)) /= 7) then
    print *, "FAIL: want [7] got [", matrix(2,2), "]"
    stop 1
end if
    if ((matrix(3,3)) /= 7) then
    print *, "FAIL: want [7] got [", matrix(3,3), "]"
    stop 1
end if
    if ((matrix(4,4)) /= 1) then
    print *, "FAIL: want [1] got [", matrix(4,4), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_nested_sections_chain
