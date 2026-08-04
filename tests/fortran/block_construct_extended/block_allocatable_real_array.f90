! vybe-test: fortran/block_construct_extended/block_allocatable_real_array
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
real, allocatable :: vals(:)
allocate(vals(3))
vals = [1.0, 2.0, 3.0]
if ((int(sum(vals))) /= 6) then
    print *, "FAIL: want [6] got [", int(sum(vals)), "]"
    stop 1
end if
end block
end program t
