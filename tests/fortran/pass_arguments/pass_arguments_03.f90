! vybe-test: fortran/pass_arguments/pass_arguments_03
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
integer :: hits = 0
type::t
contains
procedure,pass(arg)::s
end type
contains
subroutine s(arg)
class(t)::arg
hits = hits + 1
end
end module m
program driver
use m
type(t) :: obj
call obj%s()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
