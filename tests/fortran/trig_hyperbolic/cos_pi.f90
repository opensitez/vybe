! vybe-test: fortran/trig_hyperbolic/cos_pi
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
if ((nint(cos(pi)*100)) /= -100) then
    print *, "FAIL: want [-100] got [", nint(cos(pi)*100), "]"
    stop 1
end if
end program t
