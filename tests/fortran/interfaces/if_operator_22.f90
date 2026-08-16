! vybe-test: fortran/interfaces/if_operator_22
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type :: vec
integer :: v = 0
end type
interface operator(+)
module procedure addv
end interface
contains
type(vec) function addv(a,b)
type(vec), intent(in) :: a,b
addv%v = a%v + b%v
end
end module m
program driver
use m
type(vec) :: x, y, z
x%v = 2
y%v = 5
z = x + y
if (z%v /= 7) then
    print *, "FAIL: want [7] got [", z%v, "]"
    stop 1
end if
end program driver
