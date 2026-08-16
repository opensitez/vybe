! vybe-test: fortran/interfaces/if_operator_minus_34
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type :: vec
integer :: v = 0
end type
interface operator(-)
module procedure subv
end interface
contains
type(vec) function subv(a,b)
type(vec), intent(in) :: a,b
subv%v = a%v - b%v
end
end module m
program driver
use m
type(vec) :: x, y, z
x%v = 9
y%v = 4
z = x - y
if (z%v /= 5) then
    print *, "FAIL: want [5] got [", z%v, "]"
    stop 1
end if
end program driver
