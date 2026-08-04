! vybe-test: fortran/module_use_extended/use_only_one_procedure_of_two
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module arith
implicit none
contains
function add2(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = a + b
end function add2
function sub2(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = a - b
end function sub2
end module arith
program t
use arith, only: add2
if ((add2(9, 4)) /= 13) then
    print *, "FAIL: want [13] got [", add2(9, 4), "]"
    stop 1
end if
end program t
