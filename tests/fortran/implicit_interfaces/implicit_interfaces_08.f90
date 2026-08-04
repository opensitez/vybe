! vybe-test: fortran/implicit_interfaces/implicit_interfaces_08
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external s
call s(1.0)
end program p
