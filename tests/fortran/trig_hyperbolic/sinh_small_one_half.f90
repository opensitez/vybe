! vybe-test: fortran/trig_hyperbolic/sinh_small_one_half
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
if ((nint(sinh(0.5)*1000)) /= 521) then
    print *, "FAIL: want [521] got [", nint(sinh(0.5)*1000), "]"
    stop 1
end if
end program t
