! vybe-test: fortran/block_construct_extended/block_derived_type_local
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
type :: Item
integer :: id
end type Item
block
type(Item) :: it
it%id = 42
if ((it%id) /= 42) then
    print *, "FAIL: want [42] got [", it%id, "]"
    stop 1
end if
end block
end program t
