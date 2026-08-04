! vybe-test: fortran/minmaxloc_val_extended/maxval_dim1_3x2_negatives
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(3,2) = reshape([-1,-5, -2,-3, -4,-6], [3,2])
integer :: col(2)
col = maxval(m, dim=1)
if ((col(1)) /= -1) then
    print *, "FAIL: want [-1] got [", col(1), "]"
    stop 1
end if
if ((col(2)) /= -3) then
    print *, "FAIL: want [-3] got [", col(2), "]"
    stop 1
end if
end program t
