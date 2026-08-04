! vybe-test: fortran/interfaces/if_assignment_23
! origin: languages/fortran/tests/fortran/test_interfaces.rs
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
