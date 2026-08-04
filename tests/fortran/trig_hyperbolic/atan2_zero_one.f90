! vybe-test: fortran/trig_hyperbolic/atan2_zero_one
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(atan2(0.0, 1.0)*1000)) /= 0) then
    print *, "FAIL: want [0] got [", nint(atan2(0.0, 1.0)*1000), "]"
    stop 1
end if
end program t
