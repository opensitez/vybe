! vybe-test: fortran/interface_blocks/interface_blocks_10
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
module stash
integer :: seen = 0
end module stash
subroutine s(x)
use stash
integer, intent(in) :: x
seen = x
end subroutine s
program t
use stash
interface
subroutine s(x)
integer, intent(in) :: x
end subroutine s
end interface
call s(6)
if (seen /= 6) then
    print *, "FAIL: want [6] got [", seen, "]"
    stop 1
end if
end program t
