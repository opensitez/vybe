! vybe-test: fortran/block_construct_extended/block_local_integer_shadows_outer
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: x = 1
block
integer :: x
x = 100
if ((x) /= 100) then
    print *, "FAIL: want [100] got [", x, "]"
    stop 1
end if
end block
if ((x) /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
end program t
