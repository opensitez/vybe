! vybe-test: fortran/module_use_extended/module_subroutine_intent_out_fill
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module fill
implicit none
contains
subroutine fill_pair(a, b)
integer, intent(out) :: a, b
a = 2
b = 5
end subroutine fill_pair
end module fill
program t
use fill
integer :: x, y
call fill_pair(x, y)
if ((x + y) /= 7) then
    print *, "FAIL: want [7] got [", x + y, "]"
    stop 1
end if
end program t
