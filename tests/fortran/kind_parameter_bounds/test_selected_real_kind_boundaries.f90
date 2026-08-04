! vybe-test: fortran/kind_parameter_bounds/test_selected_real_kind_boundaries
! origin: languages/fortran/tests/fortran/test_kind_parameter_bounds.rs

program test_kind_parameter_bounds
    integer :: p
    p = selected_real_kind(6, 37)
    if ((p) /= 8) then
    print *, "FAIL: want [8] got [", p, "]"
    stop 1
end if
end program test_kind_parameter_bounds
