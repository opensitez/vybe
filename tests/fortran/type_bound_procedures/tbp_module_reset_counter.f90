! vybe-test: fortran/type_bound_procedures/tbp_module_reset_counter
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module tallies
implicit none
type :: Tally
integer :: n = 0
contains
procedure :: reset
end type Tally
contains
subroutine reset(self)
class(Tally), intent(inout) :: self
self%n = 0
end subroutine reset
end module tallies
program driver
use tallies
type(Tally) :: t
t%n = 9
call t%reset()
if ((t%n) /= 0) then
    print *, "FAIL: want [0] got [", t%n, "]"
    stop 1
end if
end program driver