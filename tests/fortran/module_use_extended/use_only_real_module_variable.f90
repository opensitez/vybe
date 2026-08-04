! vybe-test: fortran/module_use_extended/use_only_real_module_variable
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module phys
implicit none
real :: gravity = 9.8
real :: mass = 2.0
end module phys
program t
use phys, only: gravity
if ((int(gravity)) /= 9) then
    print *, "FAIL: want [9] got [", int(gravity), "]"
    stop 1
end if
end program t
