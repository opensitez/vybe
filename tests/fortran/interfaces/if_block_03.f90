! vybe-test: fortran/interfaces/if_block_03
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface
subroutine s(x)
integer :: x
end subroutine s
end interface
end module m
subroutine s(x)
integer :: x
x = x * 2
end subroutine s
program t
use m
integer :: v
v = 21
call s(v)
if (v /= 42) then
    print *, "FAIL: want [42] got [", v, "]"
    stop 1
end if
end program t
