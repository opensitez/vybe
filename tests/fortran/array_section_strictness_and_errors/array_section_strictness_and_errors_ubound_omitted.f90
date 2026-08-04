! vybe-test: fortran/array_section_strictness_and_errors/array_section_strictness_and_errors_ubound_omitted
! origin: languages/fortran/tests/fortran/test_array_section_strictness_and_errors.rs

program array_section_strictness_and_errors_ubound_omitted
    integer :: values(1:6)
    values = (/1, 2, 3, 4, 5, 6/)
    if ((lbound(values(3:))) /= 3) then
    print *, "FAIL: want [3] got [", lbound(values(3:)), "]"
    stop 1
end if
    if ((ubound(values(3:))) /= 6) then
    print *, "FAIL: want [6] got [", ubound(values(3:)), "]"
    stop 1
end if
    if ((size(values(3:))) /= 4) then
    print *, "FAIL: want [4] got [", size(values(3:)), "]"
    stop 1
end if
end program array_section_strictness_and_errors_ubound_omitted
