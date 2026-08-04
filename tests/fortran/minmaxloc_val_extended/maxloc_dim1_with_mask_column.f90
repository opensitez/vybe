! vybe-test: fortran/minmaxloc_val_extended/maxloc_dim1_with_mask_column
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(2,3) = reshape([1, 100, 3, 4, 5, 6], [2,3])
logical :: mask(2,3) = reshape([.true., .false., .true., .true., .true., .true.], [2,3])
integer :: col(3)
col = maxloc(m, dim=1, mask=mask)
if ((col(1)) /= 2) then
    print *, "FAIL: want [2] got [", col(1), "]"
    stop 1
end if
if ((col(2)) /= 2) then
    print *, "FAIL: want [2] got [", col(2), "]"
    stop 1
end if
end program t
