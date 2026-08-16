! vybe-test: fortran/interfaces/if_assignment_23
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type :: box
integer :: v = 0
end type
interface assignment(=)
module procedure asg
end interface
contains
subroutine asg(a,b)
type(box), intent(out) :: a
integer, intent(in) :: b
a%v = b * 2
end
end module m
program driver
use m
type(box) :: bx
bx = 5
if (bx%v /= 10) then
    print *, "FAIL: want [10] got [", bx%v, "]"
    stop 1
end if
end program driver
