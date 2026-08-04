! vybe-test: fortran/module_use_extended/public_function_returns_product
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module prodmod
implicit none
contains
public :: product2
function product2(a, b) result(r)
integer, intent(in) :: a, b
integer :: r
r = a * b
end function product2
end module prodmod
program t
use prodmod
if ((product2(6, 7)) /= 42) then
    print *, "FAIL: want [42] got [", product2(6, 7), "]"
    stop 1
end if
end program t
