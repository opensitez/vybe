! vybe-test: fortran/explicit_interfaces/explicit_interfaces_05
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
module m
interface
subroutine s(x)
integer::x
end subroutine s
end interface
end module m
