! vybe-test: fortran/minmaxloc_val_extended/maxval_real_all_equal
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
real :: a(3) = [2.5, 2.5, 2.5]
if ((int(maxval(a) * 10)) /= 25) then
    print *, "FAIL: want [25] got [", int(maxval(a) * 10), "]"
    stop 1
end if
end program t
