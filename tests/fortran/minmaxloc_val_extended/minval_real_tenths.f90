! vybe-test: fortran/minmaxloc_val_extended/minval_real_tenths
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
real :: a(5) = [0.3, 0.1, 0.4, 0.2, 0.5]
if ((int(minval(a) * 10)) /= 1) then
    print *, "FAIL: want [1] got [", int(minval(a) * 10), "]"
    stop 1
end if
end program t
