! vybe-test: fortran/trig_hyperbolic/atan_one
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(atan(1.0)*1000)) /= 785) then
    print *, "FAIL: want [785] got [", nint(atan(1.0)*1000), "]"
    stop 1
end if
end program t
