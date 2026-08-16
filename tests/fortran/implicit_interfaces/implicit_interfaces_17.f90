! vybe-test: fortran/implicit_interfaces/implicit_interfaces_17
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
module stash
integer :: seen = 0
end module stash
subroutine mix(a, b, c)
use stash
integer :: a, b, c
seen = a * 100 + b * 10 + c
end subroutine mix
program p
use stash
integer :: i, j
external mix
i = 1
j = 2
call mix(i, j, i + j)
if (seen /= 123) then
    print *, "FAIL: want [123] got [", seen, "]"
    stop 1
end if
end program p
