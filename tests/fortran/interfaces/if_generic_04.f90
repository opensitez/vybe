! vybe-test: fortran/interfaces/if_generic_04
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
integer :: hits = 0
interface g
module procedure s1
end interface
contains
subroutine s1()
hits = hits + 1
end subroutine s1
end module m
program t
use m
call g()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program t
