! vybe-test: fortran/fortran2003_extended/tbp_bound_function_hypotenuse
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
type :: Legs
real :: a, b
contains
procedure :: hyp
end type Legs
contains
function hyp(self) result(h)
class(Legs), intent(in) :: self
real :: h
h = sqrt(self%a**2 + self%b**2)
end function hyp
end module m
program driver
use m
type(Legs) :: tri
tri%a = 3.0
tri%b = 4.0
if ((int(tri%hyp())) /= 5) then
    print *, "FAIL: want [5] got [", int(tri%hyp()), "]"
    stop 1
end if
end program driver
