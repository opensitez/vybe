! vybe-test: fortran/module_use_extended/module_parameters_in_area_formula
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module shapes
implicit none
real, parameter :: PI = 3.0
real, parameter :: R = 2.0
end module shapes
program t
use shapes
if ((int(PI * R * R)) /= 12) then
    print *, "FAIL: want [12] got [", int(PI * R * R), "]"
    stop 1
end if
end program t
