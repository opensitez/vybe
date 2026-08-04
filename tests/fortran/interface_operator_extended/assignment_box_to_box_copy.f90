! vybe-test: fortran/interface_operator_extended/assignment_box_to_box_copy
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gcopy
implicit none
type :: Box
integer :: v
end type Box
interface assignment(=)
module procedure copy_box
end interface
contains
subroutine copy_box(dest, src)
type(Box), intent(out) :: dest
type(Box), intent(in) :: src
dest%v = src%v + 1
end subroutine copy_box
end module gcopy
program t
use gcopy
type(Box) :: a, b
a%v = 10
b = a
if ((b%v) /= 11) then
    print *, "FAIL: want [11] got [", b%v, "]"
    stop 1
end if
end program t
