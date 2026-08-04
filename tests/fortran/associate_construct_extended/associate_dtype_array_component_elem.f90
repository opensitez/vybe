! vybe-test: fortran/associate_construct_extended/associate_dtype_array_component_elem
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
type :: Bag
integer :: items(3)
end type Bag
type(Bag) :: b
b%items = [5, 10, 15]
associate (second => b%items(2))
if ((second) /= 10) then
    print *, "FAIL: want [10] got [", second, "]"
    stop 1
end if
end associate
end program t
