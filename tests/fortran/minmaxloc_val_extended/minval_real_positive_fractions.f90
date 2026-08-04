! vybe-test: fortran/minmaxloc_val_extended/minval_real_positive_fractions
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
real :: a(4) = [0.5, 2.5, 1.5, 3.5]
if ((int(minval(a) * 10)) /= 5) then
    print *, "FAIL: want [5] got [", int(minval(a) * 10), "]"
    stop 1
end if
end program t
