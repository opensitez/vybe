! vybe-test: fortran/module_use_extended/private_procedure_public_facade
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module facade
implicit none
private
public :: facade_double
contains
function inner_double(n) result(r)
integer, intent(in) :: n
integer :: r
r = n * 2
end function inner_double
function facade_double(n) result(r)
integer, intent(in) :: n
integer :: r
r = inner_double(n)
end function facade_double
end module facade
program t
use facade
if ((facade_double(11)) /= 22) then
    print *, "FAIL: want [22] got [", facade_double(11), "]"
    stop 1
end if
end program t
