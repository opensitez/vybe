! vybe-test: fortran/block_construct_extended/block_local_integer_from_outer_expr
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: base = 6
block
integer :: scaled
scaled = base * 3
if ((scaled) /= 18) then
    print *, "FAIL: want [18] got [", scaled, "]"
    stop 1
end if
end block
end program t
