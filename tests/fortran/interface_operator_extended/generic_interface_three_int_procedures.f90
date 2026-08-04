! vybe-test: fortran/interface_operator_extended/generic_interface_three_int_procedures
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gthree
implicit none
interface pick
module procedure pick_a, pick_b, pick_c
end interface
contains
function pick_a(v) result(r)
integer, intent(in) :: v
integer :: r
r = v
end function pick_a
function pick_b(v) result(r)
integer, intent(in) :: v
integer :: r
r = v + 1
end function pick_b
function pick_c(v) result(r)
integer, intent(in) :: v
integer :: r
r = v + 2
end function pick_c
end module gthree
program t
use gthree
if ((pick(1)) /= 1) then
    print *, "FAIL: want [1] got [", pick(1), "]"
    stop 1
end if
end program t
