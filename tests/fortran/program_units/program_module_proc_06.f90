! vybe-test: fortran/program_units/program_module_proc_06
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: hits = 0
contains
subroutine s()
hits = hits + 1
end subroutine s
end module m
program t
use m
call s()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program t
