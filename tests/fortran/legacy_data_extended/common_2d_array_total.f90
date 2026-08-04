! vybe-test: fortran/legacy_data_extended/common_2d_array_total
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: grid(2, 3)
common /grid/ grid
grid = 1
if ((sum(grid)) /= 6) then
    print *, "FAIL: want [6] got [", sum(grid), "]"
    stop 1
end if
end program t
