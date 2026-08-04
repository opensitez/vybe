! vybe-test: fortran/trig_hyperbolic/sin_array_elemental
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real, parameter :: pi = acos(-1.0)
real :: angles(2)
real :: ys(2)
angles = [0.0, pi/6.0]
ys = sin(angles)
if ((nint(ys(1)*1000)) /= 0) then
    print *, "FAIL: want [0] got [", nint(ys(1)*1000), "]"
    stop 1
end if
if ((nint(ys(2)*1000)) /= 500) then
    print *, "FAIL: want [500] got [", nint(ys(2)*1000), "]"
    stop 1
end if
end program t
