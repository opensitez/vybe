! vybe-test: fortran/fortran2003_extended/deferred_triangle_area_runtime
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module geom
implicit none
type, abstract :: Figure
contains
procedure(area_iface), deferred :: area
end type Figure
abstract interface
function area_iface(self) result(a)
import Figure
class(Figure), intent(in) :: self
real :: a
end function area_iface
end interface
type, extends(Figure) :: Tri
real :: base, height
contains
procedure :: area => tri_area
end type Tri
contains
function tri_area(self) result(a)
class(Tri), intent(in) :: self
real :: a
a = 0.5 * self%base * self%height
end function tri_area
end module geom
program t
use geom
type(Tri) :: t
t%base = 6.0
t%height = 4.0
if ((int(t%area())) /= 12) then
    print *, "FAIL: want [12] got [", int(t%area()), "]"
    stop 1
end if
end program t
