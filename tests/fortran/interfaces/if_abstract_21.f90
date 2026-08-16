! vybe-test: fortran/interfaces/if_abstract_21
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
abstract interface
subroutine s(x)
integer::x
end
end interface
end module m
program t
use m
procedure(s), pointer :: p
integer :: v
p => impl
v = 4
call p(v)
if (v /= 8) then
    print *, "FAIL: want [8] got [", v, "]"
    stop 1
end if
contains
subroutine impl(x)
integer :: x
x = x * 2
end subroutine impl
end program t
