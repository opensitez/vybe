! vybe-test: fortran/implicit_interfaces/implicit_interfaces_12
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external f
real :: x, y
x = f(y, 2)
end program p
