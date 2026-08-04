! vybe-test: fortran/interface_operator_extended/module_interface_assignment_chain
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gchain
implicit none
type :: Cell
integer :: v
end type Cell
interface assignment(=)
module procedure set_cell
end interface
contains
subroutine set_cell(dest, src)
type(Cell), intent(out) :: dest
integer, intent(in) :: src
dest%v = src * 10
end subroutine set_cell
end module gchain
program t
use gchain
type(Cell) :: c
c = 6
if ((c%v) /= 60) then
    print *, "FAIL: want [60] got [", c%v, "]"
    stop 1
end if
end program t
