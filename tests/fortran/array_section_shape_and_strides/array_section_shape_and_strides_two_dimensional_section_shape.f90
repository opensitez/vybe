! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_two_dimensional_section_shape
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_two_dimensional_section_shape
    integer :: matrix(4,5)
    matrix = reshape((/ (i, i = 1, 20) /), (/4,5/))
    if ((lbound(matrix(2:3,1:3),1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(matrix(2:3,1:3),1), "]"
    stop 1
end if
    if ((ubound(matrix(2:3,1:3),1)) /= 2) then
    print *, "FAIL: want [2] got [", ubound(matrix(2:3,1:3),1), "]"
    stop 1
end if
    if ((lbound(matrix(2:3,1:3),2)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(matrix(2:3,1:3),2), "]"
    stop 1
end if
    if ((ubound(matrix(2:3,1:3),2)) /= 3) then
    print *, "FAIL: want [3] got [", ubound(matrix(2:3,1:3),2), "]"
    stop 1
end if
    if ((sum(matrix(2:3,1:3))) /= 39) then
    print *, "FAIL: want [39] got [", sum(matrix(2:3,1:3)), "]"
    stop 1
end if
end program array_section_shape_and_strides_two_dimensional_section_shape
