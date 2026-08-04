! vybe-test: fortran/interface_operator_extended/operator_eq_boxes_not_equal
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module geq2
implicit none
type :: Box
integer :: v
end type Box
interface operator(==)
module procedure eq_box
end interface
contains
function eq_box(a, b) result(r)
type(Box), intent(in) :: a, b
logical :: r
r = a%v == b%v
end function eq_box
end module geq2
program t
use geq2
type(Box) :: a, b
a%v = 5
b%v = 6
if ((a == b) .neqv. .false.) then
    print *, "FAIL: want [false] got [", a == b, "]"
    stop 1
end if
end program t
