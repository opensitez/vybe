! vybe-test: fortran/implicit_interfaces/implicit_interfaces_05
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
subroutine caller()
external s
call s()
end subroutine caller
