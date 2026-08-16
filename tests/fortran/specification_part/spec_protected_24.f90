! vybe-test: fortran/specification_part/spec_protected_24
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
integer, protected :: x = 1
contains
subroutine bump()
x = x + 6
end subroutine bump
end module m
program t
use m
implicit none
if (x /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
call bump()
if (x /= 7) then
    print *, "FAIL: want [7] got [", x, "]"
    stop 1
end if
end program t
