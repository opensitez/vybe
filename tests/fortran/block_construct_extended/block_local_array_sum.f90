! vybe-test: fortran/block_construct_extended/block_local_array_sum
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer :: data(5)
data = [10, 20, 30, 40, 50]
if ((sum(data)) /= 150) then
    print *, "FAIL: want [150] got [", sum(data), "]"
    stop 1
end if
end block
end program t
