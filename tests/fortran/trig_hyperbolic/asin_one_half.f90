! vybe-test: fortran/trig_hyperbolic/asin_one_half
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(asin(0.5)*1000)) /= 524) then
    print *, "FAIL: want [524] got [", nint(asin(0.5)*1000), "]"
    stop 1
end if
end program t
