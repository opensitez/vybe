! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_stride_and_mask_pair
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_stride_and_mask_pair
    integer :: values(1:9)
    integer :: selected_sum
    integer :: selected_count
    values = (/ (i, i = 1, 9) /)
    selected_sum = sum(values(1:9:2), values(1:9:2) > 5)
    selected_count = count(values(1:9:2) > 5)
    if ((selected_sum) /= 7) then
    print *, "FAIL: want [7] got [", selected_sum, "]"
    stop 1
end if
    if ((selected_count) /= 1) then
    print *, "FAIL: want [1] got [", selected_count, "]"
    stop 1
end if
    if ((size(values(1:9:2))) /= 5) then
    print *, "FAIL: want [5] got [", size(values(1:9:2)), "]"
    stop 1
end if
end program array_section_shape_and_strides_stride_and_mask_pair
