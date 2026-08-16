! vybe-test: fortran/specification_part/spec_abstract_interface
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
abstract interface
 integer function f(x)
  integer, intent(in) :: x
 end function f
end interface
end module m
program t
use m
implicit none
procedure(f), pointer :: p
p => twice
if (p(6) /= 12) then
    print *, "FAIL: want [12] got [", p(6), "]"
    stop 1
end if
contains
integer function twice(x)
integer, intent(in) :: x
twice = x * 2
end function twice
end program t
