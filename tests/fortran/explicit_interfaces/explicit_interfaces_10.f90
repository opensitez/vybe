! vybe-test: fortran/explicit_interfaces/explicit_interfaces_10
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s(a)
use stash
integer, intent(in) :: a
seen = a
end subroutine s
program t
use stash
interface
subroutine s(a)
integer, intent(in) :: a
end subroutine s
end interface
call s(6)
if (seen /= 6) then
    print *, "FAIL: want [6] got [", seen, "]"
    stop 1
end if
end program t
