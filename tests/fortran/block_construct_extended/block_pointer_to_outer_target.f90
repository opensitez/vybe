! vybe-test: fortran/block_construct_extended/block_pointer_to_outer_target
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer, target :: host = 42
block
integer, pointer :: view
view => host
if ((view) /= 42) then
    print *, "FAIL: want [42] got [", view, "]"
    stop 1
end if
end block
end program t
