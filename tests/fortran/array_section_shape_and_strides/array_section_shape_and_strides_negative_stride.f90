! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_negative_stride
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_negative_stride
    integer :: values(1:9)
    values = (/ (i, i = 1, 9) /)
    if ((lbound(values(9:1:-2),1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(values(9:1:-2),1), "]"
    stop 1
end if
    if ((ubound(values(9:1:-2),1)) /= 5) then
    print *, "FAIL: want [5] got [", ubound(values(9:1:-2),1), "]"
    stop 1
end if
    if ((size(values(9:1:-2))) /= 5) then
    print *, "FAIL: want [5] got [", size(values(9:1:-2)), "]"
    stop 1
end if
    if ((sum(values(9:1:-2))) /= 25) then
    print *, "FAIL: want [25] got [", sum(values(9:1:-2)), "]"
    stop 1
end if
end program array_section_shape_and_strides_negative_stride
