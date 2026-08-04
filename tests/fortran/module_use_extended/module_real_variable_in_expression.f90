! vybe-test: fortran/module_use_extended/module_real_variable_in_expression
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module reals
implicit none
real :: rate = 2.5
end module reals
program t
use reals
if ((int(rate * 4.0)) /= 10) then
    print *, "FAIL: want [10] got [", int(rate * 4.0), "]"
    stop 1
end if
end program t
