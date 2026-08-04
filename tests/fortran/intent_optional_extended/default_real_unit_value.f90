! vybe-test: fortran/intent_optional_extended/default_real_unit_value
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((with_unit(4.0)) /= 4) then
    print *, "FAIL: want [4] got [", with_unit(4.0), "]"
    stop 1
end if
if ((with_unit(4.0, 2.0)) /= 8) then
    print *, "FAIL: want [8] got [", with_unit(4.0, 2.0), "]"
    stop 1
end if
contains
real function with_unit(v, u)
real, intent(in) :: v
real, intent(in), optional :: u
real :: use_u
if (present(u)) then
use_u = u
else
use_u = 1.0
end if
with_unit = v * use_u
end function with_unit
end program t
