! vybe-test: fortran/module_use_extended/use_only_parameter_constant
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module limits
implicit none
integer, parameter :: MAX_N = 50
integer, parameter :: MIN_N = 1
end module limits
program t
use limits, only: MAX_N
if ((MAX_N) /= 50) then
    print *, "FAIL: want [50] got [", MAX_N, "]"
    stop 1
end if
end program t
