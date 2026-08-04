! vybe-test: fortran/block_construct_extended/block_derived_type_with_array_field
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
type :: Bundle
integer :: data(3)
end type Bundle
block
type(Bundle) :: b
b%data = [2, 4, 6]
if ((sum(b%data)) /= 12) then
    print *, "FAIL: want [12] got [", sum(b%data), "]"
    stop 1
end if
end block
end program t
