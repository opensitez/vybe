! vybe-test: fortran/interface_operator_extended/operator_unary_plus_identity_box
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gplus
implicit none
type :: Box
integer :: v
end type Box
interface operator(+)
module procedure id_box
end interface
contains
function id_box(a) result(b)
type(Box), intent(in) :: a
type(Box) :: b
b%v = +a%v
end function id_box
end module gplus
program t
use gplus
type(Box) :: x, y
x%v = 8
y = +x
if ((y%v) /= 8) then
    print *, "FAIL: want [8] got [", y%v, "]"
    stop 1
end if
end program t
