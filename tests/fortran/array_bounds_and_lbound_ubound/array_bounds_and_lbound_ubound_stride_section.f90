! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_stride_section
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_stride_section
    integer :: values(1:9)
    if ((lbound(values(1:9:2), 1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(values(1:9:2), 1), "]"
    stop 1
end if
    if ((ubound(values(1:9:2), 1)) /= 5) then
    print *, "FAIL: want [5] got [", ubound(values(1:9:2), 1), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_stride_section
