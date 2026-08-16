! vybe-test: fortran/implicit_interfaces/implicit_interfaces_16
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine s(a)
use stash
integer :: a(3)
seen = sum(a)
end subroutine s
program p
use stash
external s
integer, dimension(3) :: arr
arr = [1, 2, 3]
call s(arr)
if (seen /= 6) then
    print *, "FAIL: want [6] got [", seen, "]"
    stop 1
end if
end program p
