! vybe-test: fortran/allocation/allocate_allocatable_component_runtime
! origin: languages/fortran/tests/fortran/test_allocation.rs
type :: t
integer :: x = 11
integer, allocatable :: b(:)
end type t
program t
type(t), allocatable :: obj
allocate(obj)
allocate(obj%b(3))
obj%b = [2, 4, 6]
if ((obj%x) /= 11) then
    print *, "FAIL: want [11] got [", obj%x, "]"
    stop 1
end if
if ((sum(obj%b)) /= 12) then
    print *, "FAIL: want [12] got [", sum(obj%b), "]"
    stop 1
end if
deallocate(obj%b)
deallocate(obj)
end program t
