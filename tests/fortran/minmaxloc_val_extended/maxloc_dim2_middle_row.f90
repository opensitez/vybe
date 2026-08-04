! vybe-test: fortran/minmaxloc_val_extended/maxloc_dim2_middle_row
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: m(3,3) = reshape([1,9,2, 8,3,7, 4,6,5], [3,3])
integer :: row(3)
row = maxloc(m, dim=2)
if ((row(2)) /= 1) then
    print *, "FAIL: want [1] got [", row(2), "]"
    stop 1
end if
end program t
