! vybe-test: fortran/program_units/program_submodule_style_07
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
integer :: hits = 0
interface
module subroutine s()
end subroutine s
end interface
end module m
submodule (m) msub
contains
module subroutine s()
hits = hits + 1
end subroutine s
end submodule msub
program t
use m
call s()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program t
