! vybe-test: fortran/block_construct_extended/block_modify_outer_integer
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: total = 0
block
integer :: addend
addend = 7
total = total + addend
end block
if ((total) /= 7) then
    print *, "FAIL: want [7] got [", total, "]"
    stop 1
end if
end program t
