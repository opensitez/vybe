! vybe-test: fortran/interface_operator_extended/module_interface_operator_plus_int_wrap
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module giwrap
implicit none
type :: Wrap
integer :: v
end type Wrap
interface operator(+)
module procedure add_wrap
end interface
contains
function add_wrap(a, b) result(c)
type(Wrap), intent(in) :: a, b
type(Wrap) :: c
c%v = a%v + b%v
end function add_wrap
end module giwrap
program t
use giwrap
type(Wrap) :: a, b, c
a%v = 100
b%v = 23
c = a + b
if ((c%v) /= 123) then
    print *, "FAIL: want [123] got [", c%v, "]"
    stop 1
end if
end program t
