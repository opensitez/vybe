! vybe-test: fortran/block_construct_extended/block_local_integer_uninitialized_set
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer :: n
n = 7
if ((n) /= 7) then
    print *, "FAIL: want [7] got [", n, "]"
    stop 1
end if
end block
end program t
