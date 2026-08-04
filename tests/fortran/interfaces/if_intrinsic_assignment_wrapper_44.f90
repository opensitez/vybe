! vybe-test: fortran/interfaces/if_intrinsic_assignment_wrapper_44
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface assignment(=)
module procedure asg
end interface
contains
subroutine asg(a, b)
integer, intent(out) :: a
real, intent(in) :: b
a = int(b)
end subroutine asg
end module m
