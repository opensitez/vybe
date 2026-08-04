! vybe-test: fortran/pointer_alloc_extended/alloc_2d_assign_literal_matrix
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: grid(:,:)
grid = reshape([1, 2, 3, 4], [2, 2])
if ((grid(2, 1)) /= 3) then
    print *, "FAIL: want [3] got [", grid(2, 1), "]"
    stop 1
end if
if ((sum(grid)) /= 10) then
    print *, "FAIL: want [10] got [", sum(grid), "]"
    stop 1
end if
end program t
