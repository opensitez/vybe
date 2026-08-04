! vybe-test: fortran/minmaxloc_val_extended/minval_mask_skip_first_and_last
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(7) = [100, 5, 20, 35, 50, 65, 200]
logical :: mask(7) = [.false., .true., .true., .true., .true., .true., .false.]
if ((minval(a, mask=mask)) /= 5) then
    print *, "FAIL: want [5] got [", minval(a, mask=mask), "]"
    stop 1
end if
end program t
