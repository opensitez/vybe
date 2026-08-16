! vybe-test: fortran/kind_parameter_bounds/test_kind_parameter_bounds_compute_integer_range
! origin: languages/fortran/tests/fortran/test_kind_parameter_bounds.rs

program test_kind_parameter_bounds
    integer :: small
    integer :: medium
    small = selected_int_kind(4)
    medium = selected_int_kind(8)
    if ((small) /= 2) then
    print *, "FAIL: want [2] got [", small, "]"
    stop 1
end if
    if ((medium) /= 4) then
    print *, "FAIL: want [4] got [", medium, "]"
    stop 1
end if
end program test_kind_parameter_bounds
