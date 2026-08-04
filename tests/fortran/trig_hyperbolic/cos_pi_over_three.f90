! vybe-test: fortran/trig_hyperbolic/cos_pi_over_three
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
if ((nint(cos(pi/3.0)*100)) /= 50) then
    print *, "FAIL: want [50] got [", nint(cos(pi/3.0)*100), "]"
    stop 1
end if
end program t
