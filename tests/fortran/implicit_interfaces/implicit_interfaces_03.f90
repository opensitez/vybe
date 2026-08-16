! vybe-test: fortran/implicit_interfaces/implicit_interfaces_03
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s(x)
use stash
integer :: x
seen = x * 3
end subroutine s
program p
use stash
external s
call s(1)
if (seen /= 3) then
    print *, "FAIL: want [3] got [", seen, "]"
    stop 1
end if
end program p
