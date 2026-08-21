! vybe-test: fortran/generic_interfaces/generic_push_subroutine_overload
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
type :: Stack
integer :: top = 0
contains
procedure :: push_i
procedure :: push_r
generic :: push => push_i, push_r
end type Stack
contains
subroutine push_i(self, v)
class(Stack), intent(inout) :: self
integer, intent(in) :: v
self%top = self%top + v
end subroutine push_i
subroutine push_r(self, v)
class(Stack), intent(inout) :: self
real, intent(in) :: v
self%top = self%top + int(v)
end subroutine push_r
end module m
program driver
use m
type(Stack) :: s
call s%push(1)
call s%push(2.0)
if ((s%top) /= 3) then
    print *, "FAIL: want [3] got [", s%top, "]"
    stop 1
end if
end program driver
