! vybe-test: fortran/block_construct_extended/block_nested_real_inner
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
real :: outer_r
outer_r = 2.0
block
real :: inner_r
inner_r = outer_r * 3.0
if ((int(inner_r)) /= 6) then
    print *, "FAIL: want [6] got [", int(inner_r), "]"
    stop 1
end if
end block
end block
end program t
