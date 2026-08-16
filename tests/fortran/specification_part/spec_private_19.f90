! vybe-test: fortran/specification_part/spec_private_19
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
private
integer :: hidden = 3
integer, public :: shown = 4
public :: bump
contains
subroutine bump()
hidden = hidden + 1
shown = shown + hidden
end subroutine bump
end module m
program t
use m
implicit none
call bump()
if (shown /= 8) then
    print *, "FAIL: want [8] got [", shown, "]"
    stop 1
end if
end program t
