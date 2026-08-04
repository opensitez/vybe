! vybe-test: fortran/block_construct_extended/block_outer_unchanged_after_inner_shadow
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
real :: temperature = 20.0
block
real :: temperature
temperature = 100.0
if ((int(temperature)) /= 100) then
    print *, "FAIL: want [100] got [", int(temperature), "]"
    stop 1
end if
end block
if ((int(temperature)) /= 20) then
    print *, "FAIL: want [20] got [", int(temperature), "]"
    stop 1
end if
end program t
