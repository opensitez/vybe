! vybe-test: fortran/nopass_arguments/nopass_arguments_03
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
integer :: hits = 0
type::t
contains
procedure,nopass::s1
procedure,nopass::s2
end type
contains
subroutine s1()
hits = hits + 1
end
subroutine s2()
end
end module m
program driver
use m
type(t) :: obj
call obj%s1()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
