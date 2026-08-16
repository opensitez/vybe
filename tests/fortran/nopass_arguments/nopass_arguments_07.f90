! vybe-test: fortran/nopass_arguments/nopass_arguments_07
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
integer :: hits = 0
type::t
contains
procedure,nopass::s
end type
contains
subroutine s(x)
real::x
hits = hits + 1
end
end module m
program driver
use m
type(t) :: obj
call obj%s(1.5)
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
