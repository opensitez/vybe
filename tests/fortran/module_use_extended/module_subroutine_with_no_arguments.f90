! vybe-test: fortran/module_use_extended/module_subroutine_with_no_arguments
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module ping
implicit none
contains
subroutine ping_once()
if ((0) /= 0) then
    print *, "FAIL: want [0] got [", 0, "]"
    stop 1
end if
end subroutine ping_once
end module ping
program t
use ping
call ping_once()
end program t
