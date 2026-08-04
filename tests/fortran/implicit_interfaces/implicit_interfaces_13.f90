! vybe-test: fortran/implicit_interfaces/implicit_interfaces_13
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external f
logical :: ok
ok = f(2, 3)
end program p
