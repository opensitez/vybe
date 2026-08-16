! vybe-test: fortran/implicit_interfaces/implicit_interfaces_09
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s(x)
use stash
logical :: x
seen = merge(1, 0, x)
end subroutine s
program p
use stash
external s
call s(.true.)
if (seen /= 1) then
    print *, "FAIL: want [1] got [", seen, "]"
    stop 1
end if
end program p
