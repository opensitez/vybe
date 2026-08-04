! vybe-test: fortran/minmaxloc_val_extended/maxval_real_positive_fractions
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
real :: a(4) = [0.5, 2.5, 1.5, 3.5]
if ((int(maxval(a) * 10)) /= 35) then
    print *, "FAIL: want [35] got [", int(maxval(a) * 10), "]"
    stop 1
end if
end program t
