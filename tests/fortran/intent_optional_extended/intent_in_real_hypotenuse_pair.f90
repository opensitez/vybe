! vybe-test: fortran/intent_optional_extended/intent_in_real_hypotenuse_pair
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((rhyp(3.0, 4.0)) /= 5) then
    print *, "FAIL: want [5] got [", rhyp(3.0, 4.0), "]"
    stop 1
end if
contains
real function rhyp(a, b)
real, intent(in) :: a, b
rhyp = sqrt(a*a + b*b)
end function rhyp
end program t
