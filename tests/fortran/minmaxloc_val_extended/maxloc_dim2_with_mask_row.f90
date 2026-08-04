! vybe-test: fortran/minmaxloc_val_extended/maxloc_dim2_with_mask_row
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(3,2) = reshape([10, 1, 20, 2, 30, 3], [3,2])
logical :: mask(3,2) = reshape([.true., .false., .true., .false., .true., .false.], [3,2])
integer :: row(3)
row = maxloc(m, dim=2, mask=mask)
if ((row(1)) /= 1) then
    print *, "FAIL: want [1] got [", row(1), "]"
    stop 1
end if
if ((row(3)) /= 1) then
    print *, "FAIL: want [1] got [", row(3), "]"
    stop 1
end if
end program t
