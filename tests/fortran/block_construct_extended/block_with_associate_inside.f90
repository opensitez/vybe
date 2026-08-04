! vybe-test: fortran/block_construct_extended/block_with_associate_inside
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: x = 8
block
associate (y => x + 2)
if ((y) /= 10) then
    print *, "FAIL: want [10] got [", y, "]"
    stop 1
end if
end associate
end block
end program t
