! vybe-test: fortran/intent_optional_extended/intent_in_complex_magnitude_squared
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((int(cmag2(3.0, 4.0))) /= 25) then
    print *, "FAIL: want [25] got [", int(cmag2(3.0, 4.0)), "]"
    stop 1
end if
contains
real function cmag2(re, im)
real, intent(in) :: re, im
cmag2 = re*re + im*im
end function cmag2
end program t
