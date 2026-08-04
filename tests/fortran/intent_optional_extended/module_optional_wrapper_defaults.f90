! vybe-test: fortran/intent_optional_extended/module_optional_wrapper_defaults
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
module optwrap
implicit none
contains
function bump(x, inc) result(r)
integer, intent(in) :: x
integer, intent(in), optional :: inc
integer :: r
if (present(inc)) then
r = x + inc
else
r = x + 1
end if
end function bump
end module optwrap
program t
use optwrap
if ((bump(10)) /= 11) then
    print *, "FAIL: want [11] got [", bump(10), "]"
    stop 1
end if
if ((bump(10, 4)) /= 14) then
    print *, "FAIL: want [14] got [", bump(10, 4), "]"
    stop 1
end if
end program t
