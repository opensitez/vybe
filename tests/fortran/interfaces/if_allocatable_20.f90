! vybe-test: fortran/interfaces/if_allocatable_20
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine s(x)
integer,allocatable::x(:)
end subroutine s
