! vybe-test: fortran/intent_optional_extended/intent_out_real_quarter_pi
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
real :: x
call unit_quarter(x)
if (abs((x) - 0.25) > 1.0e-6) then
    print *, "FAIL: want [0.25] got [", x, "]"
    stop 1
end if
contains
subroutine unit_quarter(v)
real, intent(out) :: v
v = 0.25
end subroutine unit_quarter
end program t
