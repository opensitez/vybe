! vybe-test: fortran/block_construct_extended/block_nested_three_levels_sum
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
integer :: l1
l1 = 1
block
integer :: l2
l2 = l1 + 2
block
integer :: l3
l3 = l2 + 3
if ((l3) /= 6) then
    print *, "FAIL: want [6] got [", l3, "]"
    stop 1
end if
end block
end block
end block
end program t
