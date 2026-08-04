! vybe-test: fortran/intent_optional_extended/default_scale_factor_ten
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((scaled(3)) /= 30) then
    print *, "FAIL: want [30] got [", scaled(3), "]"
    stop 1
end if
if ((scaled(3, 5)) /= 15) then
    print *, "FAIL: want [15] got [", scaled(3, 5), "]"
    stop 1
end if
contains
integer function scaled(v, factor)
integer, intent(in) :: v
integer, intent(in), optional :: factor
integer :: use_f
if (present(factor)) then
use_f = factor
else
use_f = 10
end if
scaled = v * use_f
end function scaled
end program t
