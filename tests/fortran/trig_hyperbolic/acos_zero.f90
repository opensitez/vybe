! vybe-test: fortran/trig_hyperbolic/acos_zero
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(acos(0.0)*1000)) /= 1571) then
    print *, "FAIL: want [1571] got [", nint(acos(0.0)*1000), "]"
    stop 1
end if
end program t
