! vybe-test: fortran/array_section_shape_and_strides/array_section_shape_and_strides_section_shape_reassigned
! origin: languages/fortran/tests/fortran/test_array_section_shape_and_strides.rs

program array_section_shape_and_strides_section_shape_reassigned
    integer :: source(1:12)
    integer :: target(1:4)
    source = (/ (i, i = 1, 12) /)
    target = source(2:8:2)
    if ((size(target)) /= 4) then
    print *, "FAIL: want [4] got [", size(target), "]"
    stop 1
end if
    if ((sum(target)) /= 20) then
    print *, "FAIL: want [20] got [", sum(target), "]"
    stop 1
end if
    if ((target(1)) /= 2) then
    print *, "FAIL: want [2] got [", target(1), "]"
    stop 1
end if
    if ((target(4)) /= 8) then
    print *, "FAIL: want [8] got [", target(4), "]"
    stop 1
end if
end program array_section_shape_and_strides_section_shape_reassigned
