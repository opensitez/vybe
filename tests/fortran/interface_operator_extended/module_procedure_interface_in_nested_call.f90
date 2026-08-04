! vybe-test: fortran/interface_operator_extended/module_procedure_interface_in_nested_call
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gnest
implicit none
interface twice
module procedure twice_val
end interface
contains
function twice_val(x) result(r)
integer, intent(in) :: x
integer :: r
r = x * 2
end function twice_val
function quad(x) result(r)
integer, intent(in) :: x
integer :: r
r = twice(x) + twice(x)
end function quad
end module gnest
program t
use gnest
if ((quad(5)) /= 20) then
    print *, "FAIL: want [20] got [", quad(5), "]"
    stop 1
end if
end program t
