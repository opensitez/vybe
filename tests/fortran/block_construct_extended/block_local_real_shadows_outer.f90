! vybe-test: fortran/block_construct_extended/block_local_real_shadows_outer
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
real :: r = 1.5
block
real :: r
r = 9.0
if ((int(r)) /= 9) then
    print *, "FAIL: want [9] got [", int(r), "]"
    stop 1
end if
end block
if ((int(r)) /= 1) then
    print *, "FAIL: want [1] got [", int(r), "]"
    stop 1
end if
end program t
