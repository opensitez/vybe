! vybe-test: fortran/pointer_alloc_extended/deallocate_then_allocated_false
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
real, allocatable :: data(:)
allocate(data(2))
data = [3.0, 4.0]
deallocate(data)
if ((allocated(data)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(data), "]"
    stop 1
end if
end program t
