! vybe-test: fortran/interface_operator_extended/operator_unary_minus_on_signed
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gneg
implicit none
type :: Signed
integer :: v
end type Signed
interface operator(-)
module procedure negate_signed
end interface
contains
function negate_signed(a) result(b)
type(Signed), intent(in) :: a
type(Signed) :: b
b%v = -a%v
end function negate_signed
end module gneg
program t
use gneg
type(Signed) :: x, y
x%v = 12
y = -x
if ((y%v) /= -12) then
    print *, "FAIL: want [-12] got [", y%v, "]"
    stop 1
end if
end program t
