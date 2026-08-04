! vybe-test: fortran/array_locators/maxloc_dim2_first_row_column_index
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])
integer :: row_maxloc(3)
row_maxloc = maxloc(m, dim=2)
if ((row_maxloc(1)) /= 2) then
    print *, "FAIL: want [2] got [", row_maxloc(1), "]"
    stop 1
end if
end program t
