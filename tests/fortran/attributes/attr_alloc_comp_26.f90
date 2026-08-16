! vybe-test: fortran/attributes/attr_alloc_comp_26
! origin: languages/fortran/tests/fortran/test_attributes.rs
program driver
type :: t
integer, allocatable :: a(:)
end type t
type(t) :: obj
allocate(obj%a(3))
obj%a = [1, 2, 3]
if (sum(obj%a) /= 6) then
    print *, "FAIL: want [6] got [", sum(obj%a), "]"
    stop 1
end if
if (size(obj%a) /= 3) then
    print *, "FAIL: want [3] got [", size(obj%a), "]"
    stop 1
end if
deallocate(obj%a)
if (merge(1, 0, allocated(obj%a)) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, allocated(obj%a)), "]"
    stop 1
end if
end program driver
