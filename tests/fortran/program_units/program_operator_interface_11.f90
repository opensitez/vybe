! vybe-test: fortran/program_units/program_operator_interface_11
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
type :: acc
integer :: v = 0
end type
interface operator(+)
module procedure adda
end interface
contains
type(acc) function adda(a,b)
type(acc), intent(in) :: a,b
adda%v = a%v + b%v
end function adda
end module m
program driver
use m
type(acc) :: x, y, z
x%v = 3
y%v = 8
z = x + y
if (z%v /= 11) then
    print *, "FAIL: want [11] got [", z%v, "]"
    stop 1
end if
end program driver
