! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_flattened_section_sum
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_flattened_section_sum
    integer :: matrix(3,3)
    integer :: total
    matrix = reshape((/ (i, i = 1, 9) /), (/3,3/))
    total = sum(reshape(matrix(2:3,1:3), (/6/))
    if ((total) /= 33) then
    print *, "FAIL: want [33] got [", total, "]"
    stop 1
end if
    if ((size(reshape(matrix(2:3,1:3), (/6/)))) /= 6) then
    print *, "FAIL: want [6] got [", size(reshape(matrix(2:3,1:3), (/6/))), "]"
    stop 1
end if
end program array_section_shape_and_strides_flattened_section_sum
