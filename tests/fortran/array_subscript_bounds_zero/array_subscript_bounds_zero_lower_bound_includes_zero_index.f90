! vybe-test: fortran/array_subscript_bounds_zero/array_subscript_bounds_zero_lower_bound_includes_zero_index
! origin: languages/fortran/tests/fortran/test_array_subscript_bounds_zero.rs

program array_subscript_bounds_zero_lower_bound_includes_zero_index
    integer :: values(0:1)
    values = (/99, 100/)
    values(0) = values(0) + 1
    if ((values(0)) /= 100) then
    print *, "FAIL: want [100] got [", values(0), "]"
    stop 1
end if
    if ((values(1)) /= 100) then
    print *, "FAIL: want [100] got [", values(1), "]"
    stop 1
end if
end program array_subscript_bounds_zero_lower_bound_includes_zero_index
