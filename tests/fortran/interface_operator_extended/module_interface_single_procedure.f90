! vybe-test: fortran/interface_operator_extended/module_interface_single_procedure
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module giface
implicit none
interface clamp_int
module procedure clamp_value
end interface
contains
function clamp_value(v, lo, hi) result(r)
integer, intent(in) :: v, lo, hi
integer :: r
if (v < lo) then
r = lo
else if (v > hi) then
r = hi
else
r = v
end if
end function clamp_value
end module giface
program t
use giface
if ((clamp_int(15, 0, 10)) /= 10) then
    print *, "FAIL: want [10] got [", clamp_int(15, 0, 10), "]"
    stop 1
end if
end program t
