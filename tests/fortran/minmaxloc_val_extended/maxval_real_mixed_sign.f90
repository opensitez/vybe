! vybe-test: fortran/minmaxloc_val_extended/maxval_real_mixed_sign
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
real :: a(5) = [-2.0, 3.0, -1.0, 4.0, 0.0]
if ((int(maxval(a))) /= 4) then
    print *, "FAIL: want [4] got [", int(maxval(a)), "]"
    stop 1
end if
end program t
