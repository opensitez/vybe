! vybe-test: fortran/minmaxloc_val_extended/maxval_dim1_4x2_matrix
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(4,2) = reshape([10,20, 30,40, 50,60, 70,80], [4,2])
integer :: col(2)
col = maxval(m, dim=1)
if ((col(1)) /= 70) then
    print *, "FAIL: want [70] got [", col(1), "]"
    stop 1
end if
if ((col(2)) /= 80) then
    print *, "FAIL: want [80] got [", col(2), "]"
    stop 1
end if
end program t
