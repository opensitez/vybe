! vybe-test: fortran/interfaces/if_assignment_real_int_35
! origin: languages/fortran/tests/fortran/test_interfaces.rs
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
