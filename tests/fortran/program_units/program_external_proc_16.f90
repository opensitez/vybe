! vybe-test: fortran/program_units/program_external_proc_16
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: hits = 0
end module m
subroutine s()
use m
hits = hits + 1
end subroutine s
program t
use m
external s
call s()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program t
