! vybe-test: fortran/module_use_extended/module_procedure_chain_three_calls
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module chain
implicit none
contains
function step1(n) result(r)
integer, intent(in) :: n
integer :: r
r = n + 1
end function step1
function step2(n) result(r)
integer, intent(in) :: n
integer :: r
r = step1(n) + 1
end function step2
end module chain
program t
use chain
if ((step2(3)) /= 5) then
    print *, "FAIL: want [5] got [", step2(3), "]"
    stop 1
end if
end program t
