! vybe-test: fortran/block_construct_extended/block_increment_outer_counter
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: n = 10
block
integer :: delta
delta = 5
n = n + delta
end block
if ((n) /= 15) then
    print *, "FAIL: want [15] got [", n, "]"
    stop 1
end if
end program t
