! vybe-test: fortran/derived_type_oop_extended/alloc_comp_deallocate_clears_component
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs
program t
type :: Bucket
integer, allocatable :: slots(:)
end type Bucket
type(Bucket) :: b
allocate(b%slots(2))
b%slots = [4, 6]
if ((b%slots(1)) /= 4) then
    print *, "FAIL: want [4] got [", b%slots(1), "]"
    stop 1
end if
deallocate(b%slots)
if ((allocated(b%slots)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(b%slots), "]"
    stop 1
end if
end program t
