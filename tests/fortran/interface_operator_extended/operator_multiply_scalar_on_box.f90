! vybe-test: fortran/interface_operator_extended/operator_multiply_scalar_on_box
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gmul
implicit none
type :: Box
integer :: v
end type Box
interface operator(*)
module procedure mul_box
end interface
contains
function mul_box(a, b) result(c)
type(Box), intent(in) :: a, b
type(Box) :: c
c%v = a%v * b%v
end function mul_box
end module gmul
program t
use gmul
type(Box) :: a, b, c
a%v = 6
b%v = 7
c = a * b
if ((c%v) /= 42) then
    print *, "FAIL: want [42] got [", c%v, "]"
    stop 1
end if
end program t
