! vybe-test: fortran/interfaces/if_abstract_21
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
abstract interface
subroutine s(x)
integer::x
end
end interface
end module m
