! vybe-test: fortran/pointer_alloc_extended/allocated_true_immediately_after_allocate
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: buf(:)
allocate(buf(4))
if ((allocated(buf)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", allocated(buf), "]"
    stop 1
end if
deallocate(buf)
end program t
