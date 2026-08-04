! vybe-test: fortran/fortran2003_extended/alloc_comp_explicit_allocate_size
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
type :: Buffer
integer, allocatable :: slots(:)
end type Buffer
type(Buffer) :: b
allocate(b%slots(4))
b%slots = [10, 20, 30, 40]
if ((b%slots(3)) /= 30) then
    print *, "FAIL: want [30] got [", b%slots(3), "]"
    stop 1
end if
end program t
