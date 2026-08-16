! vybe-test: fortran/explicit_interfaces/explicit_interfaces_08
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s(a)
use stash
integer, optional :: a
if (present(a)) then
seen = a
else
seen = -1
end if
end subroutine s
program t
use stash
interface
subroutine s(a)
integer, optional :: a
end subroutine s
end interface
call s()
if (seen /= -1) then
    print *, "FAIL: want [-1] got [", seen, "]"
    stop 1
end if
call s(8)
if (seen /= 8) then
    print *, "FAIL: want [8] got [", seen, "]"
    stop 1
end if
end program t
