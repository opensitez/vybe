! vybe-test: fortran/generic_interfaces/gen_if_04
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface assignment(=)
module procedure asg
end interface
contains
subroutine asg(a,b)
integer::a,b
a=b
end
end module m
