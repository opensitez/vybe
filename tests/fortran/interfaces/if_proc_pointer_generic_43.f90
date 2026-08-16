! vybe-test: fortran/interfaces/if_proc_pointer_generic_43
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
abstract interface
subroutine fn(x, y)
integer, intent(in) :: x
real, intent(in) :: y
end subroutine fn
end interface
integer :: seen_x = 0
real :: seen_y = 0.0
contains
subroutine impl(x, y)
integer, intent(in) :: x
real, intent(in) :: y
seen_x = x
seen_y = y
end subroutine impl
end module m
program t
use m
procedure(fn), pointer :: p
p => impl
call p(4, 2.5)
if (seen_x /= 4) then
    print *, "FAIL: want [4] got [", seen_x, "]"
    stop 1
end if
if (abs(seen_y - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", seen_y, "]"
    stop 1
end if
end program t
