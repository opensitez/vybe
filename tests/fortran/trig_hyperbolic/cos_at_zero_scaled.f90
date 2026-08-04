! vybe-test: fortran/trig_hyperbolic/cos_at_zero_scaled
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(cos(0.0)*100)) /= 100) then
    print *, "FAIL: want [100] got [", nint(cos(0.0)*100), "]"
    stop 1
end if
end program t
