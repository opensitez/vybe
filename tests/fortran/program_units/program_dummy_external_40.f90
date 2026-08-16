! vybe-test: fortran/program_units/program_dummy_external_40
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: hits = 0
end module m
subroutine target_sub()
use m
hits = hits + 1
end subroutine target_sub
subroutine apply(f)
external f
call f()
end subroutine apply
program t
use m
external target_sub
call apply(target_sub)
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program t
