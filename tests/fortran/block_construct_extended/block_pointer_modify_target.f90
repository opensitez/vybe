! vybe-test: fortran/block_construct_extended/block_pointer_modify_target
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer, target :: val = 1
block
integer, pointer :: p
p => val
p = 99
end block
if ((val) /= 99) then
    print *, "FAIL: want [99] got [", val, "]"
    stop 1
end if
end program t
