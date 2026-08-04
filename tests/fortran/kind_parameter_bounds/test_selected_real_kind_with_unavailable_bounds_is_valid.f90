! vybe-test: fortran/kind_parameter_bounds/test_selected_real_kind_with_unavailable_bounds_is_valid
! origin: languages/fortran/tests/fortran/test_kind_parameter_bounds.rs

program test_kind_parameter_bounds
    integer :: p
    p = selected_real_kind(999, 999)
    p = selected_real_kind(6, 0)
    print *, p
end program test_kind_parameter_bounds
