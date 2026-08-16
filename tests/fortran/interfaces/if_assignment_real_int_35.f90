! vybe-test: fortran/interfaces/if_assignment_real_int_35
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type :: rbox
real :: v = 0.0
end type
interface assignment(=)
module procedure asgr
end interface
contains
subroutine asgr(a,b)
type(rbox), intent(out) :: a
integer, intent(in) :: b
a%v = real(b)
end
end module m
program driver
use m
type(rbox) :: rb
rb = 3
if (abs(rb%v - 3.0) > 1.0e-6) then
    print *, "FAIL: want [3.0] got [", rb%v, "]"
    stop 1
end if
end program driver
