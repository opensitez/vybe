! vybe-test: fortran/explicit_interfaces/explicit_interfaces_09
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s(a)
use stash
integer, value :: a
a = a + 100
seen = a
end subroutine s
program t
use stash
interface
subroutine s(a)
integer, value :: a
end subroutine s
end interface
integer :: v
v = 2
call s(v)
if (v /= 2) then
    print *, "FAIL: want [2] got [", v, "]"
    stop 1
end if
if (seen /= 102) then
    print *, "FAIL: want [102] got [", seen, "]"
    stop 1
end if
end program t
