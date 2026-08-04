! vybe-test: fortran/block_construct_extended/block_allocatable_2d_array
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer, allocatable :: m(:,:)
allocate(m(2,2))
m = reshape([1, 2, 3, 4], [2,2])
if ((m(2,1)) /= 3) then
    print *, "FAIL: want [3] got [", m(2,1), "]"
    stop 1
end if
deallocate(m)
end block
end program t
