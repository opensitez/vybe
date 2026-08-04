! vybe-test: fortran/trig_hyperbolic/tanh_one
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(tanh(1.0)*1000)) /= 762) then
    print *, "FAIL: want [762] got [", nint(tanh(1.0)*1000), "]"
    stop 1
end if
end program t
