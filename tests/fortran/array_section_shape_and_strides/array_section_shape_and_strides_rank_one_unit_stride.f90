! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_rank_one_unit_stride
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_rank_one_unit_stride
    integer :: values(10)
    values = (/ (i, i = 1, 10) /)
    if ((lbound(values(2:8),1)) /= 2) then
    print *, "FAIL: want [2] got [", lbound(values(2:8),1), "]"
    stop 1
end if
    if ((ubound(values(2:8),1)) /= 8) then
    print *, "FAIL: want [8] got [", ubound(values(2:8),1), "]"
    stop 1
end if
    if ((size(values(2:8))) /= 7) then
    print *, "FAIL: want [7] got [", size(values(2:8)), "]"
    stop 1
end if
    if ((sum(values(2:8))) /= 35) then
    print *, "FAIL: want [35] got [", sum(values(2:8)), "]"
    stop 1
end if
end program array_section_shape_and_strides_rank_one_unit_stride
