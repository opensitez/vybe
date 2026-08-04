! vybe-test: fortran/module_use_extended/private_var_exposed_via_getter
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module vault
implicit none
private
public :: read_vault
integer :: stored = 99
contains
function read_vault() result(v)
integer :: v
v = stored
end function read_vault
end module vault
program t
use vault
if ((read_vault()) /= 99) then
    print *, "FAIL: want [99] got [", read_vault(), "]"
    stop 1
end if
end program t
