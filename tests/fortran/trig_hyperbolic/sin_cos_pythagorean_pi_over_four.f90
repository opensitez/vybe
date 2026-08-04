! vybe-test: fortran/trig_hyperbolic/sin_cos_pythagorean_pi_over_four
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
real :: s, c
s = sin(pi/4.0)
c = cos(pi/4.0)
if ((nint((s*s + c*c)*100)) /= 100) then
    print *, "FAIL: want [100] got [", nint((s*s + c*c)*100), "]"
    stop 1
end if
end program t
