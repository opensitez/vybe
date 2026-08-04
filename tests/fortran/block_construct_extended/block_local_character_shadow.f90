! vybe-test: fortran/block_construct_extended/block_local_character_shadow
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
character(len=4) :: s = 'outer'
block
character(len=4) :: s
s = 'inner'
if (trim(trim(s)) /= "inner") then
    print *, "FAIL: want [inner] got [", trim(s), "]"
    stop 1
end if
end block
if (trim(trim(s)) /= "outer") then
    print *, "FAIL: want [outer] got [", trim(s), "]"
    stop 1
end if
end program t
