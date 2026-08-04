! vybe-test: fortran/generic_interfaces/gen_if_15
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface g
module procedure s1
end interface
contains
subroutine s1(a,b)
integer::a,b
end
end module m
