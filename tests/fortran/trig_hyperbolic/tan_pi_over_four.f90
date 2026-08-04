! vybe-test: fortran/trig_hyperbolic/tan_pi_over_four
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
if ((nint(tan(pi/4.0)*100)) /= 100) then
    print *, "FAIL: want [100] got [", nint(tan(pi/4.0)*100), "]"
    stop 1
end if
end program t
