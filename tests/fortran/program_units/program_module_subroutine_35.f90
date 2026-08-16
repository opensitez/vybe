! vybe-test: fortran/program_units/program_module_subroutine_35
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: calls = 0
contains
subroutine s()
calls = calls + 1
end subroutine s
end module m
program t
use m
call s()
call s()
call s()
if (calls /= 3) then
    print *, "FAIL: want [3] got [", calls, "]"
    stop 1
end if
end program t
