! vybe-test: fortran/program_units/program_dummy_proc_17
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: hits = 0
end module m
subroutine bump()
use m
hits = hits + 2
end subroutine bump
subroutine apply(f)
external f
call f()
end subroutine apply
program t
use m
external bump
call apply(bump)
if (hits /= 2) then
    print *, "FAIL: want [2] got [", hits, "]"
    stop 1
end if
end program t
