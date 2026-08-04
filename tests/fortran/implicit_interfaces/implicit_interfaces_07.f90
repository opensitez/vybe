! vybe-test: fortran/implicit_interfaces/implicit_interfaces_07
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external s
call s('a')
end program p
