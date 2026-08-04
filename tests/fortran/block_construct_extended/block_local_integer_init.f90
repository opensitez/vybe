! vybe-test: fortran/block_construct_extended/block_local_integer_init
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: outer = 10
block
integer :: inner
inner = outer + 5
if ((inner) /= 15) then
    print *, "FAIL: want [15] got [", inner, "]"
    stop 1
end if
end block
end program t
