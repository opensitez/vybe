! vybe-test: fortran/minmaxloc_val_extended/maxval_mask_real_values
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
real :: a(5) = [1.1, 2.2, 3.3, 4.4, 5.5]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
if ((int(maxval(a, mask=mask) * 10)) /= 55) then
    print *, "FAIL: want [55] got [", int(maxval(a, mask=mask) * 10), "]"
    stop 1
end if
end program t
