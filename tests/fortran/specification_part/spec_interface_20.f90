! vybe-test: fortran/specification_part/spec_interface_20
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
interface
 subroutine s(x)
  integer :: x
 end subroutine s
end interface
end module m
subroutine s(x)
implicit none
integer :: x
x = x + 10
end subroutine s
program t
use m
implicit none
integer :: v
v = 1
call s(v)
if (v /= 11) then
    print *, "FAIL: want [11] got [", v, "]"
    stop 1
end if
end program t
