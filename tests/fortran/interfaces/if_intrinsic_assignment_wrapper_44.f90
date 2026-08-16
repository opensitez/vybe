! vybe-test: fortran/interfaces/if_intrinsic_assignment_wrapper_44
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type :: ibox
integer :: v = 0
end type
interface assignment(=)
module procedure asg
end interface
contains
subroutine asg(a, b)
type(ibox), intent(out) :: a
real, intent(in) :: b
a%v = int(b)
end subroutine asg
end module m
program driver
use m
type(ibox) :: ib
ib = 3.9
if (ib%v /= 3) then
    print *, "FAIL: want [3] got [", ib%v, "]"
    stop 1
end if
end program driver
