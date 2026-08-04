! vybe-test: fortran/block_construct_extended/block_local_parameter
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer, parameter :: max = 100
if ((max) /= 100) then
    print *, "FAIL: want [100] got [", max, "]"
    stop 1
end if
end block
end program t
