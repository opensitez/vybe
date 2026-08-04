! vybe-test: fortran/trig_hyperbolic/atan2_sqrt_three_one
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(atan2(sqrt(3.0), 1.0)*1000)) /= 1047) then
    print *, "FAIL: want [1047] got [", nint(atan2(sqrt(3.0), 1.0)*1000), "]"
    stop 1
end if
end program t
