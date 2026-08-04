! vybe-test: fortran/module_use_extended/public_constants_both_visible
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module constpair
implicit none
integer, public, parameter :: LOW = 2
integer, public, parameter :: HIGH = 9
end module constpair
program t
use constpair
if ((LOW + HIGH) /= 11) then
    print *, "FAIL: want [11] got [", LOW + HIGH, "]"
    stop 1
end if
end program t
