! vybe-test: fortran/module_use_extended/module_two_public_vars_summed
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module pairvals
implicit none
integer :: left = 6
integer :: right = 8
end module pairvals
program t
use pairvals
if ((left + right) /= 14) then
    print *, "FAIL: want [14] got [", left + right, "]"
    stop 1
end if
end program t
