! vybe-test: fortran/implicit_interfaces/implicit_interfaces_04
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s(x)
use stash
integer :: x
seen = x + 4
end subroutine s
program p
use stash
external s
integer :: x
x=1
call s(x)
if (seen /= 5) then
    print *, "FAIL: want [5] got [", seen, "]"
    stop 1
end if
end program p
