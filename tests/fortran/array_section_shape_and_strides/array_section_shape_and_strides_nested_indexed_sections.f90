! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_nested_indexed_sections
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_nested_indexed_sections
    integer :: matrix(6,6)
    integer :: subtotal
    matrix = reshape((/ (i, i = 1, 36) /), (/6,6/))
    subtotal = sum(matrix(2:6:2, 1:5:2))
    if ((subtotal) /= 108) then
    print *, "FAIL: want [108] got [", subtotal, "]"
    stop 1
end if
    if ((lbound(matrix(2:6:2, 1:5:2),1)) /= 2) then
    print *, "FAIL: want [2] got [", lbound(matrix(2:6:2, 1:5:2),1), "]"
    stop 1
end if
    if ((ubound(matrix(2:6:2, 1:5:2),1)) /= 6) then
    print *, "FAIL: want [6] got [", ubound(matrix(2:6:2, 1:5:2),1), "]"
    stop 1
end if
    if ((lbound(matrix(2:6:2, 1:5:2),2)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(matrix(2:6:2, 1:5:2),2), "]"
    stop 1
end if
    if ((ubound(matrix(2:6:2, 1:5:2),2)) /= 5) then
    print *, "FAIL: want [5] got [", ubound(matrix(2:6:2, 1:5:2),2), "]"
    stop 1
end if
end program array_section_shape_and_strides_nested_indexed_sections
