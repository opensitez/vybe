! vybe-test: fortran/intent_optional_extended/module_intent_out_pair_via_call
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
module outp
implicit none
contains
subroutine split10(a, b)
integer, intent(out) :: a, b
a = 4
b = 6
end subroutine split10
end module outp
program t
use outp
integer :: x, y
call split10(x, y)
if ((x + y) /= 10) then
    print *, "FAIL: want [10] got [", x + y, "]"
    stop 1
end if
end program t
