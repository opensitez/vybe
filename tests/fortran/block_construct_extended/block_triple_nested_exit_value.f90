! vybe-test: fortran/block_construct_extended/block_triple_nested_exit_value
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer :: a
a = 1
block
integer :: b
b = a + 1
block
integer :: c
c = b + 1
if ((c) /= 3) then
    print *, "FAIL: want [3] got [", c, "]"
    stop 1
end if
end block
end block
end block
end program t
