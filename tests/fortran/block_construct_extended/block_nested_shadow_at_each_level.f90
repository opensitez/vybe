! vybe-test: fortran/block_construct_extended/block_nested_shadow_at_each_level
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: v = 0
block
integer :: v
v = 10
block
integer :: v
v = 20
if ((v) /= 20) then
    print *, "FAIL: want [20] got [", v, "]"
    stop 1
end if
end block
if ((v) /= 10) then
    print *, "FAIL: want [10] got [", v, "]"
    stop 1
end if
end block
if ((v) /= 0) then
    print *, "FAIL: want [0] got [", v, "]"
    stop 1
end if
end program t
