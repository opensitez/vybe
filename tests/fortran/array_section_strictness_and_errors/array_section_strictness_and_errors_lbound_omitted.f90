! vybe-test: fortran/array_section_strictness_and_errors/array_section_strictness_and_errors_lbound_omitted
! origin: languages/fortran/tests/fortran/test_array_section_strictness_and_errors.rs

program array_section_strictness_and_errors_lbound_omitted
    integer :: values(0:4)
    values = (/1, 2, 3, 4, 5/)
    if ((lbound(values(:3),1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(values(:3),1), "]"
    stop 1
end if
    if ((ubound(values(:3),1)) /= 4) then
    print *, "FAIL: want [4] got [", ubound(values(:3),1), "]"
    stop 1
end if
    if ((size(values(:3))) /= 4) then
    print *, "FAIL: want [4] got [", size(values(:3)), "]"
    stop 1
end if
end program array_section_strictness_and_errors_lbound_omitted
