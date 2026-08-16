! vybe-test: fortran/pass_arguments/pass_arguments_08
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
integer :: hits = 0
type::t
contains
procedure,pass::a
procedure,pass::b
end type
contains
subroutine a(this)
class(t)::this
hits = hits + 1
end
subroutine b(this)
class(t)::this
end
end module m
program driver
use m
type(t) :: obj
call obj%a()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
