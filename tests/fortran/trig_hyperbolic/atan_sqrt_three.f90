! vybe-test: fortran/trig_hyperbolic/atan_sqrt_three
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(atan(sqrt(3.0))*1000)) /= 1047) then
    print *, "FAIL: want [1047] got [", nint(atan(sqrt(3.0))*1000), "]"
    stop 1
end if
end program t
