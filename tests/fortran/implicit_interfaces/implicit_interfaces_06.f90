! vybe-test: fortran/implicit_interfaces/implicit_interfaces_06
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external f
real :: x
x = f()
end program p
