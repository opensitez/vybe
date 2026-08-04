! vybe-test: fortran/trig_hyperbolic/acos_one_scaled
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(acos(1.0)*1000)) /= 0) then
    print *, "FAIL: want [0] got [", nint(acos(1.0)*1000), "]"
    stop 1
end if
end program t
