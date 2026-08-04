! vybe-test: fortran/module_use_extended/module_generic_max_mixed_kinds
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module maxmod
implicit none
interface my_max
module procedure max_int, max_real
end interface
contains
function max_int(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = max(a, b)
end function max_int
function max_real(a, b) result(r)
real, intent(in) :: a, b
real :: r
r = max(a, b)
end function max_real
end module maxmod
program t
use maxmod
if ((my_max(2, 9)) /= 9) then
    print *, "FAIL: want [9] got [", my_max(2, 9), "]"
    stop 1
end if
if ((int(my_max(1.5, 3.7))) /= 3) then
    print *, "FAIL: want [3] got [", int(my_max(1.5, 3.7)), "]"
    stop 1
end if
end program t
