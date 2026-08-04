! vybe-test: fortran/block_construct_extended/block_local_2d_array_access
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer :: grid(2,2)
grid = reshape([1, 2, 3, 4], [2,2])
if ((grid(2,2)) /= 4) then
    print *, "FAIL: want [4] got [", grid(2,2), "]"
    stop 1
end if
end block
end program t
