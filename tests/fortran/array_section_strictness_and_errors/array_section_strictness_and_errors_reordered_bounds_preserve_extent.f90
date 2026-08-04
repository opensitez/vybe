! vybe-test: fortran/array_section_strictness_and_errors/array_section_strictness_and_errors_reordered_bounds_preserve_extent
! origin: languages/fortran/tests/fortran/test_array_section_strictness_and_errors.rs

program array_section_strictness_and_errors_reordered_bounds_preserve_extent
    integer :: values(1:6)
    values = (/1, 2, 3, 4, 5, 6/)
    if ((values(6:2:-2)(1)) /= 6) then
    print *, "FAIL: want [6] got [", values(6:2:-2)(1), "]"
    stop 1
end if
    if ((values(6:2:-2)(2)) /= 2) then
    print *, "FAIL: want [2] got [", values(6:2:-2)(2), "]"
    stop 1
end if
    if ((size(values(6:2:-2))) /= 3) then
    print *, "FAIL: want [3] got [", size(values(6:2:-2)), "]"
    stop 1
end if
end program array_section_strictness_and_errors_reordered_bounds_preserve_extent
