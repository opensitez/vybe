! vybe-test: fortran/pointer_alloc_extended/alloc_derived_field_via_assignment
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
type :: Bag
integer, allocatable :: items(:)
end type Bag
type(Bag) :: box
box%items = [2, 4, 6]
if ((box%items(2)) /= 4) then
    print *, "FAIL: want [4] got [", box%items(2), "]"
    stop 1
end if
if ((sum(box%items)) /= 12) then
    print *, "FAIL: want [12] got [", sum(box%items), "]"
    stop 1
end if
end program t
