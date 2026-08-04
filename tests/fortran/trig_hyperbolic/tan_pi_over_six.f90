! vybe-test: fortran/trig_hyperbolic/tan_pi_over_six
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
if ((nint(tan(pi/6.0)*100)) /= 58) then
    print *, "FAIL: want [58] got [", nint(tan(pi/6.0)*100), "]"
    stop 1
end if
end program t
