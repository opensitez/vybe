! vybe-test: fortran/interfaces/if_bind_c_optional_42
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
integer :: seen = 0
contains
subroutine s(x) bind(c, name='s_bind')
integer, intent(in) :: x
seen = x
end subroutine s
end module m
program t
use m
call s(6)
if (seen /= 6) then
    print *, "FAIL: want [6] got [", seen, "]"
    stop 1
end if
end program t
