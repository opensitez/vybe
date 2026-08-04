! vybe-test: fortran/module_use_extended/rename_parameter_alias
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module units
implicit none
real, parameter :: METERS_PER_MILE = 1609.34
end module units
program t
use units, mile_m => METERS_PER_MILE
if ((int(mile_m)) /= 1609) then
    print *, "FAIL: want [1609] got [", int(mile_m), "]"
    stop 1
end if
end program t
