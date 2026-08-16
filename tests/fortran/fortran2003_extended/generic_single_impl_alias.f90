! vybe-test: fortran/fortran2003_extended/generic_single_impl_alias
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
type :: Box
integer :: w = 4
contains
procedure :: area_impl
generic :: area => area_impl
end type Box
contains
integer function area_impl(self) result(a)
class(Box), intent(in) :: self
a = self%w * self%w
end function area_impl
end module m
program driver
use m
type(Box) :: b
if ((b%area()) /= 16) then
    print *, "FAIL: want [16] got [", b%area(), "]"
    stop 1
end if
end program driver
