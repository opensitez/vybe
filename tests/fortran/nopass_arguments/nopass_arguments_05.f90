! vybe-test: fortran/nopass_arguments/nopass_arguments_05
! origin: languages/fortran/tests/fortran/test_nopass_arguments.rs
module m
integer :: hits = 0
type::t
contains
procedure,nopass::mk
end type
contains
function mk() result(r)
integer :: r
r=1
hits = hits + 1
end
end module m
program driver
use m
type(t) :: obj
integer :: probe
probe = obj%mk()
if (probe /= probe) then
    stop 1
end if
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
