! vybe-test: fortran/intent_optional_extended/intent_in_double_precision_diff
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((int(dsub(5.5d0, 2.0d0))) /= 3) then
    print *, "FAIL: want [3] got [", int(dsub(5.5d0, 2.0d0)), "]"
    stop 1
end if
contains
double precision function dsub(a, b)
double precision, intent(in) :: a, b
dsub = a - b
end function dsub
end program t
