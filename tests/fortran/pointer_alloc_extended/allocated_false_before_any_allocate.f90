! vybe-test: fortran/pointer_alloc_extended/allocated_false_before_any_allocate
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: buf(:)
if ((allocated(buf)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(buf), "]"
    stop 1
end if
end program t
