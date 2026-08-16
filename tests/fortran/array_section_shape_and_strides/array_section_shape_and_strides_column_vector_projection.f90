! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_column_vector_projection
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_column_vector_projection
    integer :: matrix(3,4)
    integer :: projected_sum
    matrix = reshape((/ (i, i = 1, 12) /), (/3,4/))
    projected_sum = sum(matrix(:,3))
    if ((projected_sum) /= 24) then
    print *, "FAIL: want [24] got [", projected_sum, "]"
    stop 1
end if
    if ((lbound(matrix(:,3),1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(matrix(:,3),1), "]"
    stop 1
end if
    if ((ubound(matrix(:,3),1)) /= 3) then
    print *, "FAIL: want [3] got [", ubound(matrix(:,3),1), "]"
    stop 1
end if
    if ((matrix(1,3) + matrix(3,3)) /= 16) then
    print *, "FAIL: want [16] got [", matrix(1,3) + matrix(3,3), "]"
    stop 1
end if
end program array_section_shape_and_strides_column_vector_projection
