! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_nested_indexed_sections
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_nested_indexed_sections
    integer :: matrix(6,6)
    integer :: subtotal
    matrix = reshape((/ (i, i = 1, 36) /), (/6,6/))
    subtotal = sum(matrix(2:6:2, 1:5:2))
    if ((subtotal) /= 144) then
    print *, "FAIL: want [144] got [", subtotal, "]"
    stop 1
end if
    if ((lbound(matrix(2:6:2, 1:5:2),1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(matrix(2:6:2, 1:5:2),1), "]"
    stop 1
end if
    if ((ubound(matrix(2:6:2, 1:5:2),1)) /= 3) then
    print *, "FAIL: want [3] got [", ubound(matrix(2:6:2, 1:5:2),1), "]"
    stop 1
end if
    if ((lbound(matrix(2:6:2, 1:5:2),2)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(matrix(2:6:2, 1:5:2),2), "]"
    stop 1
end if
    if ((ubound(matrix(2:6:2, 1:5:2),2)) /= 3) then
    print *, "FAIL: want [3] got [", ubound(matrix(2:6:2, 1:5:2),2), "]"
    stop 1
end if
end program array_section_shape_and_strides_nested_indexed_sections
