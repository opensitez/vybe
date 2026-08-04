! vybe-test: fortran/implicit_interfaces/implicit_interfaces_14
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external f
complex :: z
z = f(1.0, 2.0)
end program p
