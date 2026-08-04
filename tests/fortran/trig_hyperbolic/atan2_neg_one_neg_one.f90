! vybe-test: fortran/trig_hyperbolic/atan2_neg_one_neg_one
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(atan2(-1.0, -1.0)*1000)) /= -2356) then
    print *, "FAIL: want [-2356] got [", nint(atan2(-1.0, -1.0)*1000), "]"
    stop 1
end if
end program t
