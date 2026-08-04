! vybe-test: fortran/interface_operator_extended/assignment_real_pair_to_point
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gpt
implicit none
type :: Point
real :: x, y
end type Point
interface assignment(=)
module procedure pair_to_point
end interface
contains
subroutine pair_to_point(dest, src)
type(Point), intent(out) :: dest
real, intent(in) :: src(2)
dest%x = src(1)
dest%y = src(2)
end subroutine pair_to_point
end module gpt
program t
use gpt
type(Point) :: p
real :: vals(2)
vals = [3.0, 4.0]
p = vals
if ((int(p%x)) /= 3) then
    print *, "FAIL: want [3] got [", int(p%x), "]"
    stop 1
end if
if ((int(p%y)) /= 4) then
    print *, "FAIL: want [4] got [", int(p%y), "]"
    stop 1
end if
end program t
