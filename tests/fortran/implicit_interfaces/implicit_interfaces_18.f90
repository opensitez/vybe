! vybe-test: fortran/implicit_interfaces/implicit_interfaces_18
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
subroutine outer()
external s
call s(1)
end subroutine outer
