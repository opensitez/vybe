! vybe-test: fortran/interfaces/if_pass_name_36
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type::t
integer::v=0
contains
procedure,pass(self)::s
end type
contains
subroutine s(self)
class(t)::self
self%v = self%v + 3
end
end module m
program driver
use m
type(t)::obj
call obj%s()
if (obj%v /= 3) then
    print *, "FAIL: want [3] got [", obj%v, "]"
    stop 1
end if
end program driver
