! vybe-test: fortran/oop/oop_encap_14
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
integer :: hits = 0
type::t
private
integer::x
contains
procedure::setx
end type t
contains
subroutine setx(this,v)
class(t)::this
integer::v
this%x=v
hits = hits + 1
end
end module m
program driver
use m
type(t) :: obj
call obj%setx(3)
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
