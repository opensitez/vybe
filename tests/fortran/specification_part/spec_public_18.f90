! vybe-test: fortran/specification_part/spec_public_18
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
public :: x
integer :: x = 5
end module m
program t
use m
implicit none
if (x /= 5) then
    print *, "FAIL: want [5] got [", x, "]"
    stop 1
end if
x = 6
if (x /= 6) then
    print *, "FAIL: want [6] got [", x, "]"
    stop 1
end if
end program t
