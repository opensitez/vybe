! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_matrix_triangle_view
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_matrix_triangle_view
    integer :: matrix(4,4)
    matrix = reshape((/ (i, i = 1, 16) /), (/4,4/))
    if ((size(matrix(1:4:1,1:4:1),1)) /= 3) then
    print *, "FAIL: want [3] got [", size(matrix(1:4:1,1:4:1),1), "]"
    stop 1
end if
    if ((sum(matrix(2:4:1,2:4:1))) /= 85) then
    print *, "FAIL: want [85] got [", sum(matrix(2:4:1,2:4:1)), "]"
    stop 1
end if
    if ((matrix(2:4:1,2:4:1)(2,2)) /= 10) then
    print *, "FAIL: want [10] got [", matrix(2:4:1,2:4:1)(2,2), "]"
    stop 1
end if
    if ((matrix(2:4:1,2:4:1)(1,3)) /= 12) then
    print *, "FAIL: want [12] got [", matrix(2:4:1,2:4:1)(1,3), "]"
    stop 1
end if
end program array_section_shape_and_strides_matrix_triangle_view
