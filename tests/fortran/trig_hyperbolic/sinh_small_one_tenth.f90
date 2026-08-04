! vybe-test: fortran/trig_hyperbolic/sinh_small_one_tenth
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(sinh(0.1)*1000)) /= 100) then
    print *, "FAIL: want [100] got [", nint(sinh(0.1)*1000), "]"
    stop 1
end if
end program t
