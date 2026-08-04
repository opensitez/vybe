! vybe-test: fortran/intent_optional_extended/intent_out_complex_components
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
real :: re, im
call unit_imag(re, im)
if ((int(re)) /= 0) then
    print *, "FAIL: want [0] got [", int(re), "]"
    stop 1
end if
if ((int(im)) /= 1) then
    print *, "FAIL: want [1] got [", int(im), "]"
    stop 1
end if
contains
subroutine unit_imag(r, i)
real, intent(out) :: r, i
r = 0.0
i = 1.0
end subroutine unit_imag
end program t
