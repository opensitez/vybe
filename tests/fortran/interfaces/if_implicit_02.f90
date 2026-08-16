! vybe-test: fortran/interfaces/if_implicit_02
! origin: languages/fortran/tests/fortran/test_interfaces.rs
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
