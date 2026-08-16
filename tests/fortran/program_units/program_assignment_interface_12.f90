! vybe-test: fortran/program_units/program_assignment_interface_12
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
type :: box
integer :: v = 0
end type
interface assignment(=)
module procedure assigni
end interface
contains
subroutine assigni(a,b)
type(box), intent(out) :: a
integer, intent(in) :: b
a%v = b + 1
end subroutine assigni
end module m
program driver
use m
type(box) :: bx
bx = 4
if (bx%v /= 5) then
    print *, "FAIL: want [5] got [", bx%v, "]"
    stop 1
end if
end program driver
