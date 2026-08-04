! vybe-test: fortran/trig_hyperbolic/cos_pi_over_two
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
if ((nint(cos(pi/2.0)*100)) /= 0) then
    print *, "FAIL: want [0] got [", nint(cos(pi/2.0)*100), "]"
    stop 1
end if
end program t
