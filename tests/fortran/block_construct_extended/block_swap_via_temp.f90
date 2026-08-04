! vybe-test: fortran/block_construct_extended/block_swap_via_temp
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: a = 3, b = 9
block
integer :: tmp
tmp = a
a = b
b = tmp
end block
if ((a) /= 9) then
    print *, "FAIL: want [9] got [", a, "]"
    stop 1
end if
if ((b) /= 3) then
    print *, "FAIL: want [3] got [", b, "]"
    stop 1
end if
end program t
