! vybe-test: fortran/kind_parameter_bounds/test_kind_parameters_are_reusable_in_declarations
! origin: languages/fortran/tests/fortran/test_kind_parameter_bounds.rs

program test_kind_parameter_bounds
    integer, parameter :: ik = selected_int_kind(9)
    real, parameter :: rk = selected_real_kind(15)
    integer(kind=ik) :: i
    real(kind=rk) :: x
    i = 7
    x = 2.5
    print *, kind(i)
    print *, kind(x)
end program test_kind_parameter_bounds
