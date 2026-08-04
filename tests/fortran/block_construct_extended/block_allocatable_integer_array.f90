! vybe-test: fortran/block_construct_extended/block_allocatable_integer_array
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer, allocatable :: buf(:)
allocate(buf(4))
buf = [1, 2, 3, 4]
if ((sum(buf)) /= 10) then
    print *, "FAIL: want [10] got [", sum(buf), "]"
    stop 1
end if
deallocate(buf)
end block
end program t
