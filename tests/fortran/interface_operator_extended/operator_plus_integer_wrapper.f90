! vybe-test: fortran/interface_operator_extended/operator_plus_integer_wrapper
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gwrap
implicit none
type :: Box
integer :: v
end type Box
interface operator(+)
module procedure add_box
end interface
contains
function add_box(a, b) result(c)
type(Box), intent(in) :: a, b
type(Box) :: c
c%v = a%v + b%v
end function add_box
end module gwrap
program t
use gwrap
type(Box) :: x, y, z
x%v = 10
y%v = 15
z = x + y
if ((z%v) /= 25) then
    print *, "FAIL: want [25] got [", z%v, "]"
    stop 1
end if
end program t
