! vybe-test: fortran/interface_operator_extended/operator_binary_minus_point_diff
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gpdiff
implicit none
type :: Point
integer :: x, y
end type Point
interface operator(-)
module procedure sub_point
end interface
contains
function sub_point(a, b) result(c)
type(Point), intent(in) :: a, b
type(Point) :: c
c%x = a%x - b%x
c%y = a%y - b%y
end function sub_point
end module gpdiff
program t
use gpdiff
type(Point) :: a, b, c
a%x = 9; a%y = 7
b%x = 4; b%y = 2
c = a - b
if ((c%x) /= 5) then
    print *, "FAIL: want [5] got [", c%x, "]"
    stop 1
end if
if ((c%y) /= 5) then
    print *, "FAIL: want [5] got [", c%y, "]"
    stop 1
end if
end program t
