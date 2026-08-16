! vybe-test: fortran/pass_arguments/pass_arguments_05
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
integer :: hits = 0
type::t
contains
procedure,pass::show
end type
contains
subroutine show(this)
class(t)::this
print *,1
hits = hits + 1
end
end module m
program driver
use m
type(t) :: obj
call obj%show()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
