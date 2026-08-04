! vybe-test: fortran/intent_optional_extended/optional_real_tolerance_default_zero
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((near(5.0, 5.0)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", near(5.0, 5.0), "]"
    stop 1
end if
if ((near(5.0, 4.0, 0.5)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", near(5.0, 4.0, 0.5), "]"
    stop 1
end if
contains
logical function near(a, b, tol)
real, intent(in) :: a, b
real, intent(in), optional :: tol
real :: use_tol
if (present(tol)) then
use_tol = tol
else
use_tol = 0.0
end if
near = abs(a - b) <= use_tol
end function near
end program t
