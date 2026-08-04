! vybe-test: fortran/minmaxloc_val_extended/maxval_dim2_row1_of_3x3
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])
integer :: row(3)
row = maxval(m, dim=2)
if ((row(1)) /= 9) then
    print *, "FAIL: want [9] got [", row(1), "]"
    stop 1
end if
if ((row(2)) /= 8) then
    print *, "FAIL: want [8] got [", row(2), "]"
    stop 1
end if
if ((row(3)) /= 6) then
    print *, "FAIL: want [6] got [", row(3), "]"
    stop 1
end if
end program t
