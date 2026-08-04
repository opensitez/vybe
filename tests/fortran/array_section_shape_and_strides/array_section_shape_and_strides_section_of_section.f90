! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_section_of_section
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_section_of_section
    integer :: matrix(5,5)
    integer :: subtotal
    matrix = reshape((/ (i, i = 1, 25) /), (/5,5/))
    subtotal = sum(matrix(4:2:-1, 3:5))
    if ((subtotal) /= 72) then
    print *, "FAIL: want [72] got [", subtotal, "]"
    stop 1
end if
    if ((lbound(matrix(4:2:-1,3:5),1)) /= 4) then
    print *, "FAIL: want [4] got [", lbound(matrix(4:2:-1,3:5),1), "]"
    stop 1
end if
    if ((ubound(matrix(4:2:-1,3:5),1)) /= 2) then
    print *, "FAIL: want [2] got [", ubound(matrix(4:2:-1,3:5),1), "]"
    stop 1
end if
end program array_section_shape_and_strides_section_of_section
