! vybe-test: fortran/interface_operator_extended/assignment_int_to_box
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gassign
implicit none
type :: Box
integer :: v
end type Box
interface assignment(=)
module procedure int_to_box
end interface
contains
subroutine int_to_box(dest, src)
type(Box), intent(out) :: dest
integer, intent(in) :: src
dest%v = src
end subroutine int_to_box
end module gassign
program t
use gassign
type(Box) :: b
b = 42
if ((b%v) /= 42) then
    print *, "FAIL: want [42] got [", b%v, "]"
    stop 1
end if
end program t
