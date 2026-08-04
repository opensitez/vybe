! vybe-test: fortran/fortran2003_extended/tbp_accumulator_add_twice
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
type :: Acc
integer :: total = 0
contains
procedure :: bump
end type Acc
type(Acc) :: a
call a%bump(5)
call a%bump(3)
if ((a%total) /= 8) then
    print *, "FAIL: want [8] got [", a%total, "]"
    stop 1
end if
contains
subroutine bump(self, n)
class(Acc), intent(inout) :: self
integer, intent(in) :: n
self%total = self%total + n
end subroutine bump
end program t
