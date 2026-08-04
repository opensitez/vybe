! vybe-test: fortran/minmaxloc_val_extended/minval_mask_single_element
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(5) = [9, 8, 7, 6, 5]
logical :: mask(5) = [.false., .false., .true., .false., .false.]
if ((minval(a, mask=mask)) /= 7) then
    print *, "FAIL: want [7] got [", minval(a, mask=mask), "]"
    stop 1
end if
end program t
