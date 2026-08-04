! vybe-test: fortran/block_construct_extended/block_sequential_two_blocks
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer :: a
a = 3
if ((a) /= 3) then
    print *, "FAIL: want [3] got [", a, "]"
    stop 1
end if
end block
block
integer :: b
b = 7
if ((b) /= 7) then
    print *, "FAIL: want [7] got [", b, "]"
    stop 1
end if
end block
end program t
