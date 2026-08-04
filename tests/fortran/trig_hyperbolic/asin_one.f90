! vybe-test: fortran/trig_hyperbolic/asin_one
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(asin(1.0)*1000)) /= 1571) then
    print *, "FAIL: want [1571] got [", nint(asin(1.0)*1000), "]"
    stop 1
end if
end program t
