! vybe-test: fortran/implicit_interfaces/implicit_interfaces_07
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s(x)
use stash
character :: x
seen = iachar(x)
end subroutine s
program p
use stash
external s
call s('a')
if (seen /= 97) then
    print *, "FAIL: want [97] got [", seen, "]"
    stop 1
end if
end program p
