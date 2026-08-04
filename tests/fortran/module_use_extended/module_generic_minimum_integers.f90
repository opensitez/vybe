! vybe-test: fortran/module_use_extended/module_generic_minimum_integers
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module minmod
implicit none
interface my_min
module procedure min_int
end interface
contains
function min_int(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
if (a < b) then
r = a
else
r = b
end if
end function min_int
end module minmod
program t
use minmod
if ((my_min(8, 3)) /= 3) then
    print *, "FAIL: want [3] got [", my_min(8, 3), "]"
    stop 1
end if
end program t
