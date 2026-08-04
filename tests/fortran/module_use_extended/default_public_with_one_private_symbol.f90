! vybe-test: fortran/module_use_extended/default_public_with_one_private_symbol
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module mix
implicit none
integer, public :: open_val = 7
integer, private :: closed_val = 100
contains
function expose_open() result(v)
integer :: v
v = open_val
end function expose_open
end module mix
program t
use mix
if ((expose_open()) /= 7) then
    print *, "FAIL: want [7] got [", expose_open(), "]"
    stop 1
end if
end program t
