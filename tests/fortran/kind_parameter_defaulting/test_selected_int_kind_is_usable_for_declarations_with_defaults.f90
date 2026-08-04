! vybe-test: fortran/kind_parameter_defaulting/test_selected_int_kind_is_usable_for_declarations_with_defaults
! origin: languages/fortran/tests/fortran/test_kind_parameter_defaulting.rs

program test_kind_parameter_defaulting
    integer, parameter :: i4 = selected_int_kind(9)
    integer(kind=i4) :: i
    real, parameter :: r4 = selected_real_kind(5)
    real(kind=r4) :: x
    integer :: j
    real :: y
    i = 1
    x = 1.0
    j = kind(i)
    y = real(j)
    print *, kind(i)
    print *, kind(x)
end program test_kind_parameter_defaulting
