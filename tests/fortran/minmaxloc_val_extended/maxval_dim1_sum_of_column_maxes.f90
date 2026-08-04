! vybe-test: fortran/minmaxloc_val_extended/maxval_dim1_sum_of_column_maxes
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(3,3) = reshape([(i, i=1,9)], [3,3])
integer :: col(3)
col = maxval(m, dim=1)
if ((sum(col)) /= 18) then
    print *, "FAIL: want [18] got [", sum(col), "]"
    stop 1
end if
end program t
