! vybe-test: fortran/fortran2003_extended/alloc_comp_nested_container
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
type :: Inner
integer :: key = 0
end type Inner
type :: Outer
integer :: id = 1
type(Inner), allocatable :: payload(:)
end type Outer
type(Outer) :: o
allocate(o%payload(2))
o%payload(1)%key = 7
o%payload(2)%key = 3
if ((o%payload(1)%key + o%payload(2)%key) /= 10) then
    print *, "FAIL: want [10] got [", o%payload(1)%key + o%payload(2)%key, "]"
    stop 1
end if
end program t
