! vybe-test: fortran/block_construct_extended/block_local_array_fixed_size
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer :: arr(4)
arr = [1, 2, 3, 4]
if ((arr(3)) /= 3) then
    print *, "FAIL: want [3] got [", arr(3), "]"
    stop 1
end if
end block
end program t
