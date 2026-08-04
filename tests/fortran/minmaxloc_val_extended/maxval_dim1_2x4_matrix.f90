! vybe-test: fortran/minmaxloc_val_extended/maxval_dim1_2x4_matrix
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(2,4) = reshape([1,5,3,7, 2,6,4,8], [2,4])
integer :: col(4)
col = maxval(m, dim=1)
if ((col(1)) /= 2) then
    print *, "FAIL: want [2] got [", col(1), "]"
    stop 1
end if
if ((col(4)) /= 8) then
    print *, "FAIL: want [8] got [", col(4), "]"
    stop 1
end if
end program t
