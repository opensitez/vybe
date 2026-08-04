! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_matrix_section_length_probe
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_matrix_section_length_probe
    integer :: matrix(5,5)
    matrix = reshape((/ (i, i = 1, 25) /), (/5,5/))
    if ((size(matrix(2:4,3:5))) /= 6) then
    print *, "FAIL: want [6] got [", size(matrix(2:4,3:5)), "]"
    stop 1
end if
    if ((size(matrix(2:4,3:5),1)) /= 3) then
    print *, "FAIL: want [3] got [", size(matrix(2:4,3:5),1), "]"
    stop 1
end if
    if ((size(matrix(2:4,3:5),2)) /= 2) then
    print *, "FAIL: want [2] got [", size(matrix(2:4,3:5),2), "]"
    stop 1
end if
    if ((matrix(2,3)) /= 3) then
    print *, "FAIL: want [3] got [", matrix(2,3), "]"
    stop 1
end if
    if ((matrix(4,5)) /= 25) then
    print *, "FAIL: want [25] got [", matrix(4,5), "]"
    stop 1
end if
end program array_section_shape_and_strides_matrix_section_length_probe
