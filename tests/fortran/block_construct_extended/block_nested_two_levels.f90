! vybe-test: fortran/block_construct_extended/block_nested_two_levels
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: a = 1
block
integer :: b
b = a + 2
block
integer :: c
c = b + 3
if ((c) /= 6) then
    print *, "FAIL: want [6] got [", c, "]"
    stop 1
end if
end block
end block
end program t
