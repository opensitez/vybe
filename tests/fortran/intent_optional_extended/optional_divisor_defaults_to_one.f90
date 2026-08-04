! vybe-test: fortran/intent_optional_extended/optional_divisor_defaults_to_one
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((safe_div(12)) /= 12) then
    print *, "FAIL: want [12] got [", safe_div(12), "]"
    stop 1
end if
if ((safe_div(12, 4)) /= 3) then
    print *, "FAIL: want [3] got [", safe_div(12, 4), "]"
    stop 1
end if
contains
integer function safe_div(n, d)
integer, intent(in) :: n
integer, intent(in), optional :: d
integer :: use_d
if (present(d)) then
use_d = d
else
use_d = 1
end if
safe_div = n / use_d
end function safe_div
end program t
