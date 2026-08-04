! vybe-test: fortran/block_construct_extended/block_derived_type_shadow
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
type :: Node
integer :: key = 0
end type Node
type(Node) :: outer
outer%key = 1
block
type(Node) :: outer
outer%key = 99
if ((outer%key) /= 99) then
    print *, "FAIL: want [99] got [", outer%key, "]"
    stop 1
end if
end block
if ((outer%key) /= 1) then
    print *, "FAIL: want [1] got [", outer%key, "]"
    stop 1
end if
end program t
