! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_reshaped_section_compat
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_reshaped_section_compat
    integer :: matrix(2,6)
    integer :: section_sum
    matrix = reshape((/ (i, i = 1, 12) /), (/2,6/))
    section_sum = sum(reshape(matrix(1:2,2:5), (/2,2/)))
    if ((section_sum) /= 18) then
    print *, "FAIL: want [18] got [", section_sum, "]"
    stop 1
end if
    if ((lbound(matrix(1:2,2:5),1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(matrix(1:2,2:5),1), "]"
    stop 1
end if
    if ((ubound(matrix(1:2,2:5),2)) /= 4) then
    print *, "FAIL: want [4] got [", ubound(matrix(1:2,2:5),2), "]"
    stop 1
end if
end program array_section_shape_and_strides_reshaped_section_compat
