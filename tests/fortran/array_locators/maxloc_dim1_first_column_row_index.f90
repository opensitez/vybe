! vybe-test: fortran/array_locators/maxloc_dim1_first_column_row_index
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])
integer :: col_maxloc(3)
col_maxloc = maxloc(m, dim=1)
if ((col_maxloc(1)) /= 2) then
    print *, "FAIL: want [2] got [", col_maxloc(1), "]"
    stop 1
end if
end program t
