! vybe-test: fortran/interface_operator_extended/generic_abs_int_real
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gabs
implicit none
interface my_abs
module procedure abs_int, abs_real
end interface
contains
function abs_int(x) result(r)
integer, intent(in) :: x
integer :: r
if (x < 0) then
r = -x
else
r = x
end if
end function abs_int
function abs_real(x) result(r)
real, intent(in) :: x
real :: r
r = abs(x)
end function abs_real
end module gabs
program t
use gabs
if ((my_abs(-7)) /= 7) then
    print *, "FAIL: want [7] got [", my_abs(-7), "]"
    stop 1
end if
if ((int(my_abs(-7.0))) /= 7) then
    print *, "FAIL: want [7] got [", int(my_abs(-7.0)), "]"
    stop 1
end if
end program t
