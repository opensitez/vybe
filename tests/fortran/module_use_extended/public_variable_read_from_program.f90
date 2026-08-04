! vybe-test: fortran/module_use_extended/public_variable_read_from_program
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module settings
implicit none
integer, public :: level = 5
end module settings
program t
use settings
if ((level) /= 5) then
    print *, "FAIL: want [5] got [", level, "]"
    stop 1
end if
end program t
