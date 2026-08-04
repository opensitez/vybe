! vybe-test: fortran/trig_hyperbolic/acos_one_half
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(acos(0.5)*1000)) /= 1047) then
    print *, "FAIL: want [1047] got [", nint(acos(0.5)*1000), "]"
    stop 1
end if
end program t
