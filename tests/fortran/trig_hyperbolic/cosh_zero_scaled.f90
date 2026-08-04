! vybe-test: fortran/trig_hyperbolic/cosh_zero_scaled
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(cosh(0.0)*100)) /= 100) then
    print *, "FAIL: want [100] got [", nint(cosh(0.0)*100), "]"
    stop 1
end if
end program t
