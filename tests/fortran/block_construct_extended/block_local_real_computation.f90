! vybe-test: fortran/block_construct_extended/block_local_real_computation
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
real :: pi = 3.14159
block
real :: area
area = pi * 4.0
if ((int(area)) /= 12) then
    print *, "FAIL: want [12] got [", int(area), "]"
    stop 1
end if
end block
end program t
