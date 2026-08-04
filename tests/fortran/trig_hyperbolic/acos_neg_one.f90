! vybe-test: fortran/trig_hyperbolic/acos_neg_one
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(acos(-1.0)*1000)) /= 3142) then
    print *, "FAIL: want [3142] got [", nint(acos(-1.0)*1000), "]"
    stop 1
end if
end program t
