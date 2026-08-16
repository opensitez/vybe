! vybe-test: fortran/program_units/program_b_independent_15
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: hits = 0
contains
subroutine s() bind(c)
hits = hits + 1
end subroutine s
end module m
program t
use m
call s()
call s()
if (hits /= 2) then
    print *, "FAIL: want [2] got [", hits, "]"
    stop 1
end if
end program t
