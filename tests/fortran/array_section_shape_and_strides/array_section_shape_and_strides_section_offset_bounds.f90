! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_section_offset_bounds
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_section_offset_bounds
    integer :: values(-3:3)
    values = (/ (i, i = -3, 3) /)
    if ((lbound(values(-1:3),1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(values(-1:3),1), "]"
    stop 1
end if
    if ((ubound(values(-1:3),1)) /= 5) then
    print *, "FAIL: want [5] got [", ubound(values(-1:3),1), "]"
    stop 1
end if
    if ((size(values(-1:3))) /= 5) then
    print *, "FAIL: want [5] got [", size(values(-1:3)), "]"
    stop 1
end if
    if ((sum(values(-1:3))) /= 5) then
    print *, "FAIL: want [5] got [", sum(values(-1:3)), "]"
    stop 1
end if
end program array_section_shape_and_strides_section_offset_bounds
