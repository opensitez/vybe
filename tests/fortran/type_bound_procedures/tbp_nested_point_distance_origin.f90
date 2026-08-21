! vybe-test: fortran/type_bound_procedures/tbp_nested_point_distance_origin
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
type :: Coord
real :: x, y
contains
procedure :: len
end type Coord
type :: Segment
type(Coord) :: start, finish
contains
procedure :: span
end type Segment
contains
function len(self) result(d)
class(Coord), intent(in) :: self
real :: d
d = sqrt(self%x**2 + self%y**2)
end function len
function span(self) result(d)
class(Segment), intent(in) :: self
real :: d
d = sqrt((self%finish%x - self%start%x)**2 + (self%finish%y - self%start%y)**2)
end function span
end module m
program driver
use m
type(Segment) :: s
s%start%x = 0.0
s%start%y = 0.0
s%finish%x = 3.0
s%finish%y = 4.0
if ((int(s%span())) /= 5) then
    print *, "FAIL: want [5] got [", int(s%span()), "]"
    stop 1
end if
end program driver
