! vybe-test: fortran/trig_hyperbolic/acosh_roundtrip_from_cosh
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real :: x
real :: y
x = 1.0
y = acosh(cosh(x))
if ((nint(y*1000)) /= 1000) then
    print *, "FAIL: want [1000] got [", nint(y*1000), "]"
    stop 1
end if
end program t
