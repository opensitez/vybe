! vybe-test: fortran/trig_hyperbolic/tan_at_zero_scaled
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(tan(0.0)*100)) /= 0) then
    print *, "FAIL: want [0] got [", nint(tan(0.0)*100), "]"
    stop 1
end if
end program t
