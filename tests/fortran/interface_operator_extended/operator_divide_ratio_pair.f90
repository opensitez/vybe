! vybe-test: fortran/interface_operator_extended/operator_divide_ratio_pair
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gdiv
implicit none
type :: Ratio
integer :: num, den
end type Ratio
interface operator(/)
module procedure div_ratio
end interface
contains
function div_ratio(a, b) result(c)
type(Ratio), intent(in) :: a, b
type(Ratio) :: c
c%num = a%num * b%den
c%den = a%den * b%num
end function div_ratio
end module gdiv
program t
use gdiv
type(Ratio) :: a, b, c
a%num = 1; a%den = 2
b%num = 3; b%den = 4
c = a / b
if ((c%num) /= 4) then
    print *, "FAIL: want [4] got [", c%num, "]"
    stop 1
end if
if ((c%den) /= 6) then
    print *, "FAIL: want [6] got [", c%den, "]"
    stop 1
end if
end program t
