! vybe-test: fortran/intent_optional_extended/intent_inout_clamp_to_range
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: v
v = 15
call clamp(v, 0, 10)
if ((v) /= 10) then
    print *, "FAIL: want [10] got [", v, "]"
    stop 1
end if
contains
subroutine clamp(x, lo, hi)
integer, intent(inout) :: x
integer, intent(in) :: lo, hi
if (x < lo) x = lo
if (x > hi) x = hi
end subroutine clamp
end program t
