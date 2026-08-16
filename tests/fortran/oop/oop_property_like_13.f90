! vybe-test: fortran/oop/oop_property_like_13
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type::t
integer::x
contains
procedure::getx
end type t
contains
integer function getx(this)
class(t)::this
getx=this%x
end function getx
end module m
program driver
use m
type(t) :: obj
obj%x = 13
if (obj%getx() /= 13) then
    print *, "FAIL: want [13] got [", obj%getx(), "]"
    stop 1
end if
end program driver
