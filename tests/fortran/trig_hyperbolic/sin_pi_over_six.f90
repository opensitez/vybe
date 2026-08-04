! vybe-test: fortran/trig_hyperbolic/sin_pi_over_six
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
if ((nint(sin(pi/6.0)*100)) /= 50) then
    print *, "FAIL: want [50] got [", nint(sin(pi/6.0)*100), "]"
    stop 1
end if
end program t
