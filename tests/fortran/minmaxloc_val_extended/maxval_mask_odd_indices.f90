! vybe-test: fortran/minmaxloc_val_extended/maxval_mask_odd_indices
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(8) = [2, 4, 6, 8, 10, 12, 14, 16]
logical :: mask(8) = [.true., .false., .true., .false., .true., .false., .true., .false.]
if ((maxval(a, mask=mask)) /= 14) then
    print *, "FAIL: want [14] got [", maxval(a, mask=mask), "]"
    stop 1
end if
end program t
