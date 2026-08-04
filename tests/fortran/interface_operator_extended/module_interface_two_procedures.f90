! vybe-test: fortran/interface_operator_extended/module_interface_two_procedures
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gmid
implicit none
interface middle
module procedure mid_int, mid_real
end interface
contains
function mid_int(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = (a + b) / 2
end function mid_int
function mid_real(a, b) result(r)
real, intent(in) :: a, b
real :: r
r = (a + b) / 2.0
end function mid_real
end module gmid
program t
use gmid
if ((middle(3, 7)) /= 5) then
    print *, "FAIL: want [5] got [", middle(3, 7), "]"
    stop 1
end if
if ((int(middle(3.0, 7.0))) /= 5) then
    print *, "FAIL: want [5] got [", int(middle(3.0, 7.0)), "]"
    stop 1
end if
end program t
