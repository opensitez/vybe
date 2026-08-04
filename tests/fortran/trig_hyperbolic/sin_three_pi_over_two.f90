! vybe-test: fortran/trig_hyperbolic/sin_three_pi_over_two
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
if ((nint(sin(3.0*pi/2.0)*100)) /= -100) then
    print *, "FAIL: want [-100] got [", nint(sin(3.0*pi/2.0)*100), "]"
    stop 1
end if
end program t
