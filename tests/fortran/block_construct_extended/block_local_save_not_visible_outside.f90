! vybe-test: fortran/block_construct_extended/block_local_save_not_visible_outside
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer :: secret
secret = 42
if ((secret) /= 42) then
    print *, "FAIL: want [42] got [", secret, "]"
    stop 1
end if
end block
if (trim('done') /= "done") then
    print *, "FAIL: want [done] got [", 'done', "]"
    stop 1
end if
end program t
