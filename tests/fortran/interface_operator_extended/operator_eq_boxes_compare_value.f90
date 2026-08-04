! vybe-test: fortran/interface_operator_extended/operator_eq_boxes_compare_value
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module geq
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
end module geq
program t
use geq
type(Box) :: a, b
a%v = 5
b%v = 5
if ((a == b) .neqv. .true.) then
    print *, "FAIL: want [true] got [", a == b, "]"
    stop 1
end if
end program t
