! vybe-test: fortran/generic_interfaces/gen_if_13
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface assignment(=)
module procedure asgr
end interface
contains
subroutine asgr(a,b)
real::a
integer::b
a=real(b)
end
end module m
