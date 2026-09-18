! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_strided_column_slice
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_strided_column_slice
    integer :: matrix(3,6)
    matrix = reshape((/ (i, i = 1, 18) /), (/3,6/))
    if ((size(matrix(:,2:6:2),1)) /= 3) then
    print *, "FAIL: want [3] got [", size(matrix(:,2:6:2),1), "]"
    stop 1
end if
    if ((size(matrix(:,2:6:2),2)) /= 3) then
    print *, "FAIL: want [3] got [", size(matrix(:,2:6:2),2), "]"
    stop 1
end if
    if ((sum(matrix(:,2:6:2))) /= 99) then
    print *, "FAIL: want [99] got [", sum(matrix(:,2:6:2)), "]"
    stop 1
end if
    if ((matrix(2,2:6:2)(1)) /= 5) then
    print *, "FAIL: want [5] got [", matrix(2,2:6:2)(1), "]"
    stop 1
end if
    if ((matrix(2,2:6:2)(size(matrix(:,2:6:2),2))) /= 17) then
    print *, "FAIL: want [17] got [", matrix(2,2:6:2)(size(matrix(:,2:6:2),2)), "]"
    stop 1
end if
end program array_section_shape_and_strides_strided_column_slice
