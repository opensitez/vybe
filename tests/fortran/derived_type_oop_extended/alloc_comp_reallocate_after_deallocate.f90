! vybe-test: fortran/derived_type_oop_extended/alloc_comp_reallocate_after_deallocate
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs
program t
type :: Buffer
integer, allocatable :: data(:)
end type Buffer
type(Buffer) :: buf
allocate(buf%data(2))
buf%data = [1, 2]
deallocate(buf%data)
allocate(buf%data(3))
buf%data = [3, 4, 5]
if ((buf%data(3)) /= 5) then
    print *, "FAIL: want [5] got [", buf%data(3), "]"
    stop 1
end if
end program t
