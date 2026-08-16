! vybe-test: fortran/interfaces/if_nopass_14
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
integer :: hits = 0
type::t
contains
procedure,nopass::s
end type
contains
subroutine s()
hits = hits + 1
end
end module m
program driver
use m
type(t)::obj
call obj%s()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
