! vybe-test: fortran/explicit_interfaces/explicit_interfaces_01
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s()
use stash
seen = seen + 1
end subroutine s
program t
use stash
interface
subroutine s()
end subroutine s
end interface
call s()
call s()
if (seen /= 2) then
    print *, "FAIL: want [2] got [", seen, "]"
    stop 1
end if
end program t
