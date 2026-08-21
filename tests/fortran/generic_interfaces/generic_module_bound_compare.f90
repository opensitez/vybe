! vybe-test: fortran/generic_interfaces/generic_module_bound_compare
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module cmpmod
implicit none
type :: Cmp
contains
procedure :: eq_i
procedure :: eq_r
generic :: eq => eq_i, eq_r
end type Cmp
contains
logical function eq_i(self, a, b) result(r)
class(Cmp), intent(in) :: self
integer, intent(in) :: a, b
r = a == b
end function eq_i
logical function eq_r(self, a, b) result(r)
class(Cmp), intent(in) :: self
real, intent(in) :: a, b
r = abs(a - b) < 1.0e-6
end function eq_r
end module cmpmod
program t
use cmpmod
type(Cmp) :: c
if (.not. (c%eq(4, 4))) then
    print *, "FAIL: want [1] got [", c%eq(4, 4), "]"
    stop 1
end if
if (.not. (c%eq(1.0, 1.0))) then
    print *, "FAIL: want [1] got [", c%eq(1.0, 1.0), "]"
    stop 1
end if
end program t
