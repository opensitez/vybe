! vybe-test: fortran/program_units/program_binding_label_14
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: hits = 0
contains
subroutine s() bind(c, name='s_c')
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
