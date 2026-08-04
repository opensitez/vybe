! vybe-test: fortran/block_construct_extended/block_local_character_string
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
character(len=6) :: msg
msg = 'block'
if (trim(trim(msg)) /= "block") then
    print *, "FAIL: want [block] got [", trim(msg), "]"
    stop 1
end if
end block
end program t
