! vybe-test: fortran/associate_construct_extended/associate_nested_dtype_field
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
type :: Node
integer :: key
end type Node
type(Node) :: n
n%key = 42
associate (item => n)
associate (k => item%key)
if ((k) /= 42) then
    print *, "FAIL: want [42] got [", k, "]"
    stop 1
end if
end associate
end associate
end program t
