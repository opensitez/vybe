! vybe-test: fortran/module_use_extended/use_only_function_from_module
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module calc
implicit none
contains
function triple(n) result(r)
integer, intent(in) :: n
integer :: r
r = n * 3
end function triple
end module calc
program t
use calc, only: triple
if ((triple(4)) /= 12) then
    print *, "FAIL: want [12] got [", triple(4), "]"
    stop 1
end if
end program t
