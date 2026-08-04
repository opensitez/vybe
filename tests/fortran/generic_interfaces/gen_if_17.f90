! vybe-test: fortran/generic_interfaces/gen_if_17
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
abstract interface
subroutine s(x)
integer::x
end
end interface
end module m
