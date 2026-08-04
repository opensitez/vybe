! vybe-test: fortran/minmaxloc_val_extended/minval_dim1_sum_of_column_mins
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(3,3) = reshape([(i, i=1,9)], [3,3])
integer :: col(3)
col = minval(m, dim=1)
if ((sum(col)) /= 12) then
    print *, "FAIL: want [12] got [", sum(col), "]"
    stop 1
end if
end program t
