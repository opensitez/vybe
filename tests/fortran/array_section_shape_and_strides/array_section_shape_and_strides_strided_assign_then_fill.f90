! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_strided_assign_then_fill
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_strided_assign_then_fill
    integer :: values(1:20)
    integer :: section_sum
    values = 0
    values(2:20:3) = 4
    section_sum = sum(values(2:20:3))
    if ((section_sum) /= 28) then
    print *, "FAIL: want [28] got [", section_sum, "]"
    stop 1
end if
    if ((size(values(2:20:3))) /= 7) then
    print *, "FAIL: want [7] got [", size(values(2:20:3)), "]"
    stop 1
end if
    if ((values(2)) /= 4) then
    print *, "FAIL: want [4] got [", values(2), "]"
    stop 1
end if
    if ((values(20)) /= 4) then
    print *, "FAIL: want [4] got [", values(20), "]"
    stop 1
end if
end program array_section_shape_and_strides_strided_assign_then_fill
