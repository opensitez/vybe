! vybe-test: fortran/interface_blocks/interface_blocks_08
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
module stash
integer :: seen = 0
end module stash
subroutine s(x)
use stash
integer, optional :: x
if (present(x)) then
seen = x
else
seen = -1
end if
end subroutine s
program t
use stash
interface
subroutine s(x)
integer, optional :: x
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
