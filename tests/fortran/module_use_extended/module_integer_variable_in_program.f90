! vybe-test: fortran/module_use_extended/module_integer_variable_in_program
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module ints
implicit none
integer :: seed = 13
end module ints
program t
use ints
if ((seed) /= 13) then
    print *, "FAIL: want [13] got [", seed, "]"
    stop 1
end if
end program t
