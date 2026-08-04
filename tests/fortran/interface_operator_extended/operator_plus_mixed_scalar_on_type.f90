! vybe-test: fortran/interface_operator_extended/operator_plus_mixed_scalar_on_type
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gshift
implicit none
type :: Offset
integer :: delta
end type Offset
interface operator(+)
module procedure add_offset
end interface
contains
function add_offset(a, b) result(c)
type(Offset), intent(in) :: a, b
type(Offset) :: c
c%delta = a%delta + b%delta
end function add_offset
end module gshift
program t
use gshift
type(Offset) :: a, b, c
a%delta = 4
b%delta = -1
c = a + b
if ((c%delta) /= 3) then
    print *, "FAIL: want [3] got [", c%delta, "]"
    stop 1
end if
end program t
