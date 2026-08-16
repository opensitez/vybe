! vybe-test: fortran/implicit_interfaces/implicit_interfaces_01
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s()
use stash
seen = 7
end subroutine s
program p
use stash
external s
call s()
if (seen /= 7) then
    print *, "FAIL: want [7] got [", seen, "]"
    stop 1
end if
end program p
