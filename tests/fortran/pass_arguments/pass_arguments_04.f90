! vybe-test: fortran/pass_arguments/pass_arguments_04
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
integer :: hits = 0
type::t
contains
procedure,pass::s1
procedure,pass::s2
end type
contains
subroutine s1(this)
class(t)::this
hits = hits + 1
end
subroutine s2(this)
class(t)::this
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
