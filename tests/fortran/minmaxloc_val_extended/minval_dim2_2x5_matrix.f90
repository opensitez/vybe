! vybe-test: fortran/minmaxloc_val_extended/minval_dim2_2x5_matrix
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(2,5) = reshape([1,3,5,7,9, 2,4,6,8,10], [2,5])
integer :: row(2)
row = minval(m, dim=2)
if ((row(1)) /= 1) then
    print *, "FAIL: want [1] got [", row(1), "]"
    stop 1
end if
if ((row(2)) /= 2) then
    print *, "FAIL: want [2] got [", row(2), "]"
    stop 1
end if
end program t
