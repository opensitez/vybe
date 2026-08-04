! vybe-test: fortran/explicit_interfaces/explicit_interfaces_10
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
interface
subroutine s(a)
integer, intent(in) :: a
end subroutine s
end interface
