! vybe-test: fortran/trig_hyperbolic/tanh_small_one_half
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(tanh(0.5)*1000)) /= 462) then
    print *, "FAIL: want [462] got [", nint(tanh(0.5)*1000), "]"
    stop 1
end if
end program t
