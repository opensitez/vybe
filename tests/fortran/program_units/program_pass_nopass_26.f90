! vybe-test: fortran/program_units/program_pass_nopass_26
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: passed = 0
integer :: nopassed = 0
type :: t
 integer :: v = 7
contains
 procedure, pass :: s1
 procedure, nopass :: s2
end type t
contains
subroutine s1(this)
 class(t) :: this
 passed = this%v
end subroutine s1
subroutine s2()
 nopassed = 1
end subroutine s2
end module m
program driver
use m
type(t) :: obj
call obj%s1()
if (passed /= 7) then
    print *, "FAIL: want [7] got [", passed, "]"
    stop 1
end if
call obj%s2()
if (nopassed /= 1) then
    print *, "FAIL: want [1] got [", nopassed, "]"
    stop 1
end if
end program driver
