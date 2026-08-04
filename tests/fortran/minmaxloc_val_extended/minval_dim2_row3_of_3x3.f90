! vybe-test: fortran/minmaxloc_val_extended/minval_dim2_row3_of_3x3
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])
integer :: row(3)
row = minval(m, dim=2)
if ((row(3)) /= 4) then
    print *, "FAIL: want [4] got [", row(3), "]"
    stop 1
end if
end program t
