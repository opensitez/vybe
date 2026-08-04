! vybe-test: fortran/explicit_interfaces/explicit_interfaces_07
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
interface
subroutine s(a)
real::a(:)
end subroutine s
end interface
