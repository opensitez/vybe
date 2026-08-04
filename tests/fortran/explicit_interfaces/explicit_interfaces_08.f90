! vybe-test: fortran/explicit_interfaces/explicit_interfaces_08
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
interface
subroutine s(a)
integer, optional :: a
end subroutine s
end interface
