! vybe-test: fortran/interface_operator_extended/assignment_then_operator_on_type
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gmix
implicit none
type :: Acc
integer :: v
end type Acc
interface assignment(=)
module procedure set_acc
end interface
interface operator(+)
module procedure add_acc
end interface
contains
subroutine set_acc(dest, src)
type(Acc), intent(out) :: dest
integer, intent(in) :: src
dest%v = src
end subroutine set_acc
function add_acc(a, b) result(c)
type(Acc), intent(in) :: a, b
type(Acc) :: c
c%v = a%v + b%v
end function add_acc
end module gmix
program t
use gmix
type(Acc) :: x, y, z
x = 4
y = 5
z = x + y
if ((z%v) /= 9) then
    print *, "FAIL: want [9] got [", z%v, "]"
    stop 1
end if
end program t
