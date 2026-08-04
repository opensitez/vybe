! vybe-test: fortran/intent_optional_extended/optional_real_pair_weighted
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((int(wavg(2.0, 8.0))) /= 5) then
    print *, "FAIL: want [5] got [", int(wavg(2.0, 8.0)), "]"
    stop 1
end if
if ((int(wavg(2.0, 8.0, 0.25))) /= 6) then
    print *, "FAIL: want [6] got [", int(wavg(2.0, 8.0, 0.25)), "]"
    stop 1
end if
contains
real function wavg(a, b, w)
real, intent(in) :: a, b
real, intent(in), optional :: w
real :: use_w
if (present(w)) then
use_w = w
else
use_w = 0.5
end if
wavg = a * use_w + b * (1.0 - use_w)
end function wavg
end program t
