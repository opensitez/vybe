! vybe-test: fortran/program_units/program_recursive_04
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: depth = 0
integer :: deepest = 0
contains
recursive subroutine s(n)
integer :: n
depth = depth + 1
if (depth > deepest) deepest = depth
if (n > 0) call s(n - 1)
depth = depth - 1
end subroutine s
end module m
program t
use m
call s(3)
if (deepest /= 4) then
    print *, "FAIL: want [4] got [", deepest, "]"
    stop 1
end if
end program t
