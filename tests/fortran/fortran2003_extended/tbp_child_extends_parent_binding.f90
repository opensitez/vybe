! vybe-test: fortran/fortran2003_extended/tbp_child_extends_parent_binding
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
type :: Base
integer :: n = 1
contains
procedure :: twice
end type Base
type, extends(Base) :: Child
integer :: extra = 0
end type Child
contains
function twice(self) result(v)
class(Base), intent(in) :: self
integer :: v
v = self%n * 2
end function twice
end module m
program driver
use m
type(Child) :: c
if ((c%twice()) /= 2) then
    print *, "FAIL: want [2] got [", c%twice(), "]"
    stop 1
end if
end program driver
