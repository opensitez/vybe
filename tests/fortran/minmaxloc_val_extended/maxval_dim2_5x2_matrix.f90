! vybe-test: fortran/minmaxloc_val_extended/maxval_dim2_5x2_matrix
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(5,2) = reshape([11,12, 21,22, 31,32, 41,42, 51,52], [5,2])
integer :: row(5)
row = maxval(m, dim=2)
if ((row(1)) /= 12) then
    print *, "FAIL: want [12] got [", row(1), "]"
    stop 1
end if
if ((row(5)) /= 52) then
    print *, "FAIL: want [52] got [", row(5), "]"
    stop 1
end if
end program t
