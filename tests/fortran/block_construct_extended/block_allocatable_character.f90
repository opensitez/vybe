! vybe-test: fortran/block_construct_extended/block_allocatable_character
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
block
character(len=:), allocatable :: s
s = 'hello'
if ((len_trim(s)) /= 5) then
    print *, "FAIL: want [5] got [", len_trim(s), "]"
    stop 1
end if
end block
end program t
