! vybe-test: fortran/implicit_interfaces/implicit_interfaces_20
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
implicit none
integer :: x
external f
x = f(x)
end program p
