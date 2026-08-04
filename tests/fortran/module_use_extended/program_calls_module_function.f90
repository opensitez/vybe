! vybe-test: fortran/module_use_extended/program_calls_module_function
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module funcs
implicit none
contains
function halve(n) result(r)
real, intent(in) :: n
real :: r
r = n / 2.0
end function halve
end module funcs
program t
use funcs
if ((int(halve(14.0))) /= 7) then
    print *, "FAIL: want [7] got [", int(halve(14.0)), "]"
    stop 1
end if
end program t
