! vybe-test: fortran/interface_operator_extended/operator_plus_point2d_sum_components
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gpoint
implicit none
type :: Point
real :: x, y
end type Point
interface operator(+)
module procedure add_point
end interface
contains
function add_point(a, b) result(c)
type(Point), intent(in) :: a, b
type(Point) :: c
c%x = a%x + b%x
c%y = a%y + b%y
end function add_point
end module gpoint
program t
use gpoint
type(Point) :: p, q, r
p%x = 1.0; p%y = 2.0
q%x = 3.0; q%y = 4.0
r = p + q
if ((int(r%x)) /= 4) then
    print *, "FAIL: want [4] got [", int(r%x), "]"
    stop 1
end if
if ((int(r%y)) /= 6) then
    print *, "FAIL: want [6] got [", int(r%y), "]"
    stop 1
end if
end program t
