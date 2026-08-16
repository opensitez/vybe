! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_section_from_non_default_base
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_section_from_non_default_base
    integer :: values(-5:5)
    if ((lbound(values( -2:3), 1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(values( -2:3), 1), "]"
    stop 1
end if
    if ((ubound(values( -2:3), 1)) /= 6) then
    print *, "FAIL: want [6] got [", ubound(values( -2:3), 1), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_section_from_non_default_base
