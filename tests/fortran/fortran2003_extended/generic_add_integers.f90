! vybe-test: fortran/fortran2003_extended/generic_add_integers
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
type :: Adder
contains
procedure :: add_i
generic :: add => add_i
end type Adder
type(Adder) :: a
if ((a%add(3, 5)) /= 8) then
    print *, "FAIL: want [8] got [", a%add(3, 5), "]"
    stop 1
end if
contains
integer function add_i(self, x, y) result(r)
class(Adder), intent(in) :: self
integer, intent(in) :: x, y
r = x + y
end function add_i
end program t
