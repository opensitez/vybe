! vybe-test: fortran/fortran2003_extended/generic_int_show_binding
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
type :: Printer
contains
procedure :: show_i
procedure :: show_r
generic :: show => show_i, show_r
end type Printer
contains
subroutine show_i(self, v)
class(Printer), intent(in) :: self
integer, intent(in) :: v
if ((v) /= 7) then
    print *, "FAIL: want [7] got [", v, "]"
    stop 1
end if
end subroutine show_i
subroutine show_r(self, v)
class(Printer), intent(in) :: self
real, intent(in) :: v
if ((int(v)) /= 2) then
    print *, "FAIL: want [2] got [", int(v), "]"
    stop 1
end if
end subroutine show_r
end module m
program driver
use m
type(Printer) :: p
call p%show(7)
call p%show(2.5)
end program driver
