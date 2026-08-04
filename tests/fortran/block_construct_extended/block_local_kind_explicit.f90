! vybe-test: fortran/block_construct_extended/block_local_kind_explicit
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer(kind=8) :: big
big = 10000000000
if ((int(big / 1000000000)) /= 10) then
    print *, "FAIL: want [10] got [", int(big / 1000000000), "]"
    stop 1
end if
end block
end program t
