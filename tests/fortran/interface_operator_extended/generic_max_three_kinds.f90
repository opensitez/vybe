! vybe-test: fortran/interface_operator_extended/generic_max_three_kinds
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gmax
implicit none
interface pick_max
module procedure max_int, max_real, max_logical
end interface
contains
function max_int(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = max(a, b)
end function max_int
function max_real(a, b) result(r)
real, intent(in) :: a, b
real :: r
r = max(a, b)
end function max_real
function max_logical(a, b) result(r)
logical, intent(in) :: a, b
logical :: r
if (a .eqv. b) then
r = a
else
r = .true.
end if
end function max_logical
end module gmax
program t
use gmax
if ((pick_max(2, 9)) /= 9) then
    print *, "FAIL: want [9] got [", pick_max(2, 9), "]"
    stop 1
end if
if ((int(pick_max(2.0, 9.0))) /= 9) then
    print *, "FAIL: want [9] got [", int(pick_max(2.0, 9.0)), "]"
    stop 1
end if
if ((pick_max(.false., .true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", pick_max(.false., .true.), "]"
    stop 1
end if
end program t
