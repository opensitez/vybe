! vybe-test: fortran/minmaxloc_val_extended/maxval_dim2_real_2x3
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
real :: m(2,3) = reshape([1.0, 3.0, 2.0, 6.0, 4.0, 5.0], [2,3])
real :: row(2)
row = maxval(m, dim=2)
if ((int(row(1))) /= 4) then
    print *, "FAIL: want [4] got [", int(row(1)), "]"
    stop 1
end if
if ((int(row(2))) /= 6) then
    print *, "FAIL: want [6] got [", int(row(2)), "]"
    stop 1
end if
end program t
