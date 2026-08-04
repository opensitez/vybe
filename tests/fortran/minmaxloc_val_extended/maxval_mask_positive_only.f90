! vybe-test: fortran/minmaxloc_val_extended/maxval_mask_positive_only
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(6) = [-5, 3, -2, 8, -1, 4]
logical :: mask(6) = [.false., .true., .false., .true., .false., .true.]
if ((maxval(a, mask=mask)) /= 8) then
    print *, "FAIL: want [8] got [", maxval(a, mask=mask), "]"
    stop 1
end if
end program t
