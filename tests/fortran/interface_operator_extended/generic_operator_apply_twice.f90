! vybe-test: fortran/interface_operator_extended/generic_operator_apply_twice
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gtwice
implicit none
type :: Num
integer :: v
end type Num
interface operator(+)
module procedure add_num
end interface
contains
function add_num(a, b) result(c)
type(Num), intent(in) :: a, b
type(Num) :: c
c%v = a%v + b%v
end function add_num
end module gtwice
program t
use gtwice
type(Num) :: a, b, c, d
a%v = 1
b%v = 2
c%v = 3
d = a + b + c
if ((d%v) /= 6) then
    print *, "FAIL: want [6] got [", d%v, "]"
    stop 1
end if
end program t
