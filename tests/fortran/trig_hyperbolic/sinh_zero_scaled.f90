! vybe-test: fortran/trig_hyperbolic/sinh_zero_scaled
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(sinh(0.0)*100)) /= 0) then
    print *, "FAIL: want [0] got [", nint(sinh(0.0)*100), "]"
    stop 1
end if
end program t
