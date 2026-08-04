! vybe-test: fortran/trig_hyperbolic/cosh_small_one_half
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(cosh(0.5)*1000)) /= 1128) then
    print *, "FAIL: want [1128] got [", nint(cosh(0.5)*1000), "]"
    stop 1
end if
end program t
