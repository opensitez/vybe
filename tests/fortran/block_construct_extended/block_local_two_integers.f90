! vybe-test: fortran/block_construct_extended/block_local_two_integers
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer :: a, b
a = 3
b = 4
if ((a + b) /= 7) then
    print *, "FAIL: want [7] got [", a + b, "]"
    stop 1
end if
end block
end program t
