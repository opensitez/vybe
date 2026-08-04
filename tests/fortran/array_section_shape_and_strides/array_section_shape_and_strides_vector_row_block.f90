! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_vector_row_block
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_vector_row_block
    integer :: matrix(4,4)
    matrix = reshape((/ (i, i = 1, 16) /), (/4,4/))
    if ((size(matrix(2:3,2:4),1)) /= 2) then
    print *, "FAIL: want [2] got [", size(matrix(2:3,2:4),1), "]"
    stop 1
end if
    if ((size(matrix(2:3,2:4),2)) /= 3) then
    print *, "FAIL: want [3] got [", size(matrix(2:3,2:4),2), "]"
    stop 1
end if
    if ((matrix(2,2:4)(2)) /= 7) then
    print *, "FAIL: want [7] got [", matrix(2,2:4)(2), "]"
    stop 1
end if
    if ((sum(matrix(2:3,2:4))) /= 58) then
    print *, "FAIL: want [58] got [", sum(matrix(2:3,2:4)), "]"
    stop 1
end if
end program array_section_shape_and_strides_vector_row_block
