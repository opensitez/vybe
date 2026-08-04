! vybe-test: fortran/intent_optional_extended/default_cap_when_max_absent
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((capped(12)) /= 12) then
    print *, "FAIL: want [12] got [", capped(12), "]"
    stop 1
end if
if ((capped(12, 10)) /= 10) then
    print *, "FAIL: want [10] got [", capped(12, 10), "]"
    stop 1
end if
contains
integer function capped(v, lim)
integer, intent(in) :: v
integer, intent(in), optional :: lim
integer :: use_lim
if (present(lim)) then
use_lim = lim
else
use_lim = 100
end if
if (v > use_lim) then
capped = use_lim
else
capped = v
end if
end function capped
end program t
