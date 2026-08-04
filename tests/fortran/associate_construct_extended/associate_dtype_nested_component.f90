! vybe-test: fortran/associate_construct_extended/associate_dtype_nested_component
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
type :: Inner
integer :: val = 0
end type Inner
type :: Outer
type(Inner) :: core
end type Outer
type(Outer) :: o
o%core%val = 77
associate (v => o%core%val)
if ((v) /= 77) then
    print *, "FAIL: want [77] got [", v, "]"
    stop 1
end if
end associate
end program t
