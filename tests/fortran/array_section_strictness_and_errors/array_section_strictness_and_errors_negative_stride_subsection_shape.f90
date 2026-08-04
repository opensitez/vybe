! vybe-test: fortran/array_section_strictness_and_errors/array_section_strictness_and_errors_negative_stride_subsection_shape
! origin: languages/fortran/tests/fortran/test_array_section_strictness_and_errors.rs

program array_section_strictness_and_errors_negative_stride_subsection_shape
    integer :: values(1:9)
    values = (/1, 2, 3, 4, 5, 6, 7, 8, 9/)
    if ((size(values(9:1:-2))) /= 5) then
    print *, "FAIL: want [5] got [", size(values(9:1:-2)), "]"
    stop 1
end if
    if ((values(9:1:-2)(1)) /= 9) then
    print *, "FAIL: want [9] got [", values(9:1:-2)(1), "]"
    stop 1
end if
    if ((values(9:1:-2)(3)) /= 5) then
    print *, "FAIL: want [5] got [", values(9:1:-2)(3), "]"
    stop 1
end if
    if ((sum(values(9:1:-2))) /= 25) then
    print *, "FAIL: want [25] got [", sum(values(9:1:-2)), "]"
    stop 1
end if
end program array_section_strictness_and_errors_negative_stride_subsection_shape
