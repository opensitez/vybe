! vybe-test: fortran/attributes/attr_deferred_12
! origin: languages/fortran/tests/fortran/test_attributes.rs
module m
type, abstract :: t
contains
procedure(p),deferred::s
end type t
abstract interface
subroutine p(this)
import t
class(t)::this
end
end interface
type, extends(t) :: impl
integer :: v = 0
contains
procedure :: s => impl_s
end type impl
contains
subroutine impl_s(this)
class(impl)::this
this%v = 7
end subroutine impl_s
end module m
program driver
use m
type(impl) :: obj
call obj%s()
if (obj%v /= 7) then
    print *, "FAIL: want [7] got [", obj%v, "]"
    stop 1
end if
end program driver
