! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_subsection_for_sum
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_subsection_for_sum
    integer :: values(1:15)
    values = (/ (i, i = 1, 15) /)
    if ((sum(values(5:15:5))) /= 30) then
    print *, "FAIL: want [30] got [", sum(values(5:15:5)), "]"
    stop 1
end if
    if ((size(values(5:15:5))) /= 3) then
    print *, "FAIL: want [3] got [", size(values(5:15:5)), "]"
    stop 1
end if
    if ((values(5:15:5)(1)) /= 5) then
    print *, "FAIL: want [5] got [", values(5:15:5)(1), "]"
    stop 1
end if
    if ((values(5:15:5)(3)) /= 15) then
    print *, "FAIL: want [15] got [", values(5:15:5)(3), "]"
    stop 1
end if
end program array_section_shape_and_strides_subsection_for_sum
