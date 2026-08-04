! vybe-test: fortran/module_use_extended/module_interface_operator_unary_negate
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module neg
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
end module neg
program t
use neg
type(Signed) :: x, y
x%v = 7
y = -x
if ((y%v) /= -7) then
    print *, "FAIL: want [-7] got [", y%v, "]"
    stop 1
end if
end program t
