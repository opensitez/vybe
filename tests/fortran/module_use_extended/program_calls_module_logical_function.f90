! vybe-test: fortran/module_use_extended/program_calls_module_logical_function
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module checks
implicit none
contains
function is_zero(n) result(b)
integer, intent(in) :: n
logical :: b
b = n == 0
end function is_zero
end module checks
program t
use checks
if ((is_zero(0)) /= 1) then
    print *, "FAIL: want [1] got [", is_zero(0), "]"
    stop 1
end if
end program t
