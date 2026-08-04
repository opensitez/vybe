! vybe-test: fortran/implicit_interfaces/implicit_interfaces_11
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external f
integer :: x
x = f()
end program p
