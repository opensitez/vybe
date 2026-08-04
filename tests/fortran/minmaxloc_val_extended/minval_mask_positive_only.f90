! vybe-test: fortran/minmaxloc_val_extended/minval_mask_positive_only
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(6) = [-5, 3, -2, 8, -1, 4]
logical :: mask(6) = [.false., .true., .false., .true., .false., .true.]
if ((minval(a, mask=mask)) /= 3) then
    print *, "FAIL: want [3] got [", minval(a, mask=mask), "]"
    stop 1
end if
end program t
