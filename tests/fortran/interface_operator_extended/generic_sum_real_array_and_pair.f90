! vybe-test: fortran/interface_operator_extended/generic_sum_real_array_and_pair
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gsum
implicit none
interface total
module procedure total_pair, total_real2
end interface
contains
function total_pair(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = a + b
end function total_pair
function total_real2(x, y) result(r)
real, intent(in) :: x, y
real :: r
r = x + y
end function total_real2
end module gsum
program t
use gsum
if ((total(4, 5)) /= 9) then
    print *, "FAIL: want [9] got [", total(4, 5), "]"
    stop 1
end if
if ((int(total(1.5, 2.5))) /= 4) then
    print *, "FAIL: want [4] got [", int(total(1.5, 2.5)), "]"
    stop 1
end if
end program t
