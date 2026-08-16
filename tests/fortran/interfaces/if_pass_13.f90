! vybe-test: fortran/interfaces/if_pass_13
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type::t
integer::v=0
contains
procedure,pass::s
end type
contains
subroutine s(this)
class(t)::this
this%v = this%v + 4
end
end module m
program driver
use m
type(t)::obj
call obj%s()
if (obj%v /= 4) then
    print *, "FAIL: want [4] got [", obj%v, "]"
    stop 1
end if
end program driver
