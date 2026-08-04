! vybe-test: fortran/interface_operator_extended/operator_multiply_accumulate_boxes
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gacc
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
end module gacc
program t
use gacc
type(Box) :: a, b, c, d
a%v = 2
b%v = 3
c%v = 4
d = a * b * c
if ((d%v) /= 24) then
    print *, "FAIL: want [24] got [", d%v, "]"
    stop 1
end if
end program t
