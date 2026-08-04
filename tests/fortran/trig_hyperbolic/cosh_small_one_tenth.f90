! vybe-test: fortran/trig_hyperbolic/cosh_small_one_tenth
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(cosh(0.1)*1000)) /= 1005) then
    print *, "FAIL: want [1005] got [", nint(cosh(0.1)*1000), "]"
    stop 1
end if
end program t
