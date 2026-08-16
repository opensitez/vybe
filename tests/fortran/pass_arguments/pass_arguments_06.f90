! vybe-test: fortran/pass_arguments/pass_arguments_06
! origin: languages/fortran/tests/fortran/test_pass_arguments.rs
module m
integer :: hits = 0
type::t
contains
procedure,pass::get
end type
contains
integer function get(this)
class(t)::this
get=1
hits = hits + 1
end
end module m
program driver
use m
type(t) :: obj
integer :: probe
probe = obj%get()
if (probe /= probe) then
    stop 1
end if
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
