! vybe-test: fortran/implicit_interfaces/implicit_interfaces_15
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external f
character(len=10) :: text
text = f()
end program p
