! vybe-test: fortran/minmaxloc_val_extended/minval_dim2_5x2_matrix
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(5,2) = reshape([11,12, 21,22, 31,32, 41,42, 51,52], [5,2])
integer :: row(5)
row = minval(m, dim=2)
if ((row(3)) /= 21) then
    print *, "FAIL: want [21] got [", row(3), "]"
    stop 1
end if
end program t
