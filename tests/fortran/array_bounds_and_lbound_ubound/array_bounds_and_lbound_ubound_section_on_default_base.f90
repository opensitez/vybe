! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_section_on_default_base
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_section_on_default_base
    integer :: values(10)
    if ((lbound(values(3:7), 1)) /= 3) then
    print *, "FAIL: want [3] got [", lbound(values(3:7), 1), "]"
    stop 1
end if
    if ((ubound(values(3:7), 1)) /= 7) then
    print *, "FAIL: want [7] got [", ubound(values(3:7), 1), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_section_on_default_base
