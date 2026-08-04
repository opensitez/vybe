! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_step_two_stride
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_step_two_stride
    integer :: values(1:12)
    values = (/ (i, i = 1, 12) /)
    if ((lbound(values(2:10:2),1)) /= 2) then
    print *, "FAIL: want [2] got [", lbound(values(2:10:2),1), "]"
    stop 1
end if
    if ((size(values(2:10:2))) /= 5) then
    print *, "FAIL: want [5] got [", size(values(2:10:2)), "]"
    stop 1
end if
    if ((values(2:10:2)(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(2:10:2)(1), "]"
    stop 1
end if
    if ((values(2:10:2)(size(values(2:10:2)))) /= 10) then
    print *, "FAIL: want [10] got [", values(2:10:2)(size(values(2:10:2))), "]"
    stop 1
end if
    if ((sum(values(2:10:2))) /= 30) then
    print *, "FAIL: want [30] got [", sum(values(2:10:2)), "]"
    stop 1
end if
end program array_section_shape_and_strides_step_two_stride
