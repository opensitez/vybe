! vybe-test: fortran/minmaxloc_val_extended/minval_mask_odd_indices
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(8) = [2, 4, 6, 8, 10, 12, 14, 16]
logical :: mask(8) = [.true., .false., .true., .false., .true., .false., .true., .false.]
if ((minval(a, mask=mask)) /= 2) then
    print *, "FAIL: want [2] got [", minval(a, mask=mask), "]"
    stop 1
end if
end program t
