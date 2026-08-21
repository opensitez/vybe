! vybe-test: fortran/generic_interfaces/generic_add_integers
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
type :: Adder
contains
procedure :: add_i
generic :: add => add_i
end type Adder
contains
integer function add_i(self, x, y) result(r)
class(Adder), intent(in) :: self
integer, intent(in) :: x, y
r = x + y
end function add_i
end module m
program driver
use m
type(Adder) :: a
if ((a%add(3, 5)) /= 8) then
    print *, "FAIL: want [8] got [", a%add(3, 5), "]"
    stop 1
end if
end program driver
