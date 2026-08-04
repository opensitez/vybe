! vybe-test: fortran/intent_optional_extended/intent_inout_real_halve_value
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
real :: x
x = 8.0
call halve_real(x)
if ((x) /= 4) then
    print *, "FAIL: want [4] got [", x, "]"
    stop 1
end if
contains
subroutine halve_real(v)
real, intent(inout) :: v
v = v / 2.0
end subroutine halve_real
end program t
