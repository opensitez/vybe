! vybe-test: fortran/kind_parameter_bounds/test_selected_int_kind_falls_back_to_unavailable
! origin: languages/fortran/tests/fortran/test_kind_parameter_bounds.rs

program test_kind_parameter_bounds
    if ((selected_int_kind(1000)) /= -1) then
    print *, "FAIL: want [-1] got [", selected_int_kind(1000), "]"
    stop 1
end if
end program test_kind_parameter_bounds
