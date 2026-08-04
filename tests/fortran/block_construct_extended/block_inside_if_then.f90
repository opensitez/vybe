! vybe-test: fortran/block_construct_extended/block_inside_if_then
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: x = 5
if (x > 0) then
block
integer :: y
y = x * 2
if ((y) /= 10) then
    print *, "FAIL: want [10] got [", y, "]"
    stop 1
end if
end block
end if
end program t
