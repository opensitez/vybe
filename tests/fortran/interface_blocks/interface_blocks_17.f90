! vybe-test: fortran/interface_blocks/interface_blocks_17
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
module m
type :: holder
integer :: v = 0
end type
end module m
subroutine assign_wrapper(lhs, rhs)
use m
type(holder), intent(out) :: lhs
integer, intent(in) :: rhs
lhs%v = rhs * 2
end subroutine assign_wrapper
program driver
use m
interface assignment(=)
subroutine assign_wrapper(lhs, rhs)
use m
type(holder), intent(out) :: lhs
integer, intent(in) :: rhs
end subroutine assign_wrapper
end interface
type(holder) :: h
h = 6
if (h%v /= 12) then
    print *, "FAIL: want [12] got [", h%v, "]"
    stop 1
end if
end program driver
