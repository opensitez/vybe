! vybe-test: fortran/trig_hyperbolic/asinh_zero
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(asinh(0.0)*100)) /= 0) then
    print *, "FAIL: want [0] got [", nint(asinh(0.0)*100), "]"
    stop 1
end if
end program t
