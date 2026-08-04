! vybe-test: fortran/interface_operator_extended/assignment_box_from_two_integers
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gpair
implicit none
type :: PairBox
integer :: a, b
end type PairBox
interface assignment(=)
module procedure ints_to_pairbox
end interface
contains
subroutine ints_to_pairbox(dest, src)
type(PairBox), intent(out) :: dest
integer, intent(in) :: src(2)
dest%a = src(1)
dest%b = src(2)
end subroutine ints_to_pairbox
end module gpair
program t
use gpair
type(PairBox) :: p
integer :: v(2)
v = [2, 3]
p = v
if ((p%a + p%b) /= 5) then
    print *, "FAIL: want [5] got [", p%a + p%b, "]"
    stop 1
end if
end program t
