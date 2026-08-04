! vybe-test: fortran/minmaxloc_val_extended/maxval_dim2_2x5_matrix
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(2,5) = reshape([1,3,5,7,9, 2,4,6,8,10], [2,5])
integer :: row(2)
row = maxval(m, dim=2)
if ((row(1)) /= 9) then
    print *, "FAIL: want [9] got [", row(1), "]"
    stop 1
end if
if ((row(2)) /= 10) then
    print *, "FAIL: want [10] got [", row(2), "]"
    stop 1
end if
end program t
