! vybe-test: fortran/block_construct_extended/block_pointer_reassign_in_block
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer, target :: a = 5, b = 8
block
integer, pointer :: p
p => a
if ((p) /= 5) then
    print *, "FAIL: want [5] got [", p, "]"
    stop 1
end if
p => b
if ((p) /= 8) then
    print *, "FAIL: want [8] got [", p, "]"
    stop 1
end if
end block
end program t
