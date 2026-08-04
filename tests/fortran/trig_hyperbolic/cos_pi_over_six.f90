! vybe-test: fortran/trig_hyperbolic/cos_pi_over_six
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
if ((nint(cos(pi/6.0)*100)) /= 87) then
    print *, "FAIL: want [87] got [", nint(cos(pi/6.0)*100), "]"
    stop 1
end if
end program t
