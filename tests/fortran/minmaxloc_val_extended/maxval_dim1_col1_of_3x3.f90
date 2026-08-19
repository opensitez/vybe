! vybe-test: fortran/minmaxloc_val_extended/maxval_dim1_col1_of_3x3
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])
integer :: col(3)
col = maxval(m, dim=1)
if ((col(1)) /= 9) then
    print *, "FAIL: want [9] got [", col(1), "]"
    stop 1
end if
if ((col(2)) /= 8) then
    print *, "FAIL: want [8] got [", col(2), "]"
    stop 1
end if
if ((col(3)) /= 6) then
    print *, "FAIL: want [6] got [", col(3), "]"
    stop 1
end if
end program t
