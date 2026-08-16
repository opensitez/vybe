! vybe-test: fortran/implicit_interfaces/implicit_interfaces_08
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s(x)
use stash
real :: x
seen = nint(x * 2.0)
end subroutine s
program p
use stash
external s
call s(1.0)
if (seen /= 2) then
    print *, "FAIL: want [2] got [", seen, "]"
    stop 1
end if
end program p
