! vybe-test: fortran/fortran2003_extended/deferred_rect_area_runtime
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module shapes
implicit none
type, abstract :: Poly
contains
procedure(perim_iface), deferred :: perimeter
end type Poly
abstract interface
function perim_iface(self) result(p)
import Poly
class(Poly), intent(in) :: self
real :: p
end function perim_iface
end interface
type, extends(Poly) :: Rect
real :: w, h
contains
procedure :: perimeter => rect_perim
end type Rect
contains
function rect_perim(self) result(p)
class(Rect), intent(in) :: self
real :: p
p = 2.0 * (self%w + self%h)
end function rect_perim
end module shapes
program t
use shapes
type(Rect) :: r
r%w = 3.0
r%h = 4.0
if ((int(r%perimeter())) /= 14) then
    print *, "FAIL: want [14] got [", int(r%perimeter()), "]"
    stop 1
end if
end program t
