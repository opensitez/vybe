! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_section_as_lhs_rhs_same_shape
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_section_as_lhs_rhs_same_shape
    integer :: values(1:10)
    integer :: sample(2:8)
    values = (/ (i, i = 1, 10) /)
    sample = values(2:8)
    sample(2:4) = sample(2:4) + 1
    if ((sum(sample)) /= 38) then
    print *, "FAIL: want [38] got [", sum(sample), "]"
    stop 1
end if
    if ((sample(2)) /= 3) then
    print *, "FAIL: want [3] got [", sample(2), "]"
    stop 1
end if
    if ((sample(8)) /= 8) then
    print *, "FAIL: want [8] got [", sample(8), "]"
    stop 1
end if
end program array_section_shape_and_strides_section_as_lhs_rhs_same_shape
