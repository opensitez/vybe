! vybe-test: fortran/interface_blocks/interface_blocks_09
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
module stash
integer :: seen = 0
end module stash
subroutine s(x)
use stash
integer, value :: x
x = x + 100
seen = x
end subroutine s
program t
use stash
interface
subroutine s(x)
integer, value :: x
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
