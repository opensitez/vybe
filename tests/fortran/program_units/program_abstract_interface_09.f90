! vybe-test: fortran/program_units/program_abstract_interface_09
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
abstract interface
subroutine s(x)
integer :: x
end subroutine s
end interface
end module m
program t
use m
procedure(s), pointer :: p
integer :: v
p => impl
v = 5
call p(v)
if (v /= 15) then
    print *, "FAIL: want [15] got [", v, "]"
    stop 1
end if
contains
subroutine impl(x)
integer :: x
x = x * 3
end subroutine impl
end program t
