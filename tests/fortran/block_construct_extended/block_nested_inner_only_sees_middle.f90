! vybe-test: fortran/block_construct_extended/block_nested_inner_only_sees_middle
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: x = 5
block
integer :: y
y = x + 1
block
integer :: z
z = y + 2
if ((z) /= 8) then
    print *, "FAIL: want [8] got [", z, "]"
    stop 1
end if
end block
if ((y) /= 6) then
    print *, "FAIL: want [6] got [", y, "]"
    stop 1
end if
end block
end program t
