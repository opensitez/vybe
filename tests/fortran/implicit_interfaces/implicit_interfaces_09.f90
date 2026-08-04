! vybe-test: fortran/implicit_interfaces/implicit_interfaces_09
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external s
call s(.true.)
end program p
