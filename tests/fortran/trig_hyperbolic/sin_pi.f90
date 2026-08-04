! vybe-test: fortran/trig_hyperbolic/sin_pi
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
if ((nint(sin(pi)*100)) /= 0) then
    print *, "FAIL: want [0] got [", nint(sin(pi)*100), "]"
    stop 1
end if
end program t
