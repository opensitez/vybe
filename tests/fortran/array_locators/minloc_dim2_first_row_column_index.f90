! vybe-test: fortran/array_locators/minloc_dim2_first_row_column_index
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])
integer :: row_minloc(3)
row_minloc = minloc(m, dim=2)
if ((row_minloc(1)) /= 1) then
    print *, "FAIL: want [1] got [", row_minloc(1), "]"
    stop 1
end if
end program t
