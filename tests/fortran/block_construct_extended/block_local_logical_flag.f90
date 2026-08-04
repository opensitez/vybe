! vybe-test: fortran/block_construct_extended/block_local_logical_flag
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
logical :: ok
ok = .true.
if ((ok) .neqv. .true.) then
    print *, "FAIL: want [true] got [", ok, "]"
    stop 1
end if
end block
end program t
