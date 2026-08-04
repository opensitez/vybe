! vybe-test: fortran/intent_optional_extended/default_floor_when_min_absent
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((floored(-3)) /= 0) then
    print *, "FAIL: want [0] got [", floored(-3), "]"
    stop 1
end if
if ((floored(-3, 0)) /= 0) then
    print *, "FAIL: want [0] got [", floored(-3, 0), "]"
    stop 1
end if
contains
integer function floored(v, lo)
integer, intent(in) :: v
integer, intent(in), optional :: lo
integer :: use_lo
if (present(lo)) then
use_lo = lo
else
use_lo = 0
end if
if (v < use_lo) then
floored = use_lo
else
floored = v
end if
end function floored
end program t
