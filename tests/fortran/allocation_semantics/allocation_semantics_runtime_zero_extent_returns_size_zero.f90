! vybe-test: fortran/allocation_semantics/allocation_semantics_runtime_zero_extent_returns_size_zero
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program t
integer, allocatable :: a(:)
allocate(a(0))
if ((size(a)) /= 0) then
    print *, "FAIL: want [0] got [", size(a), "]"
    stop 1
end if
deallocate(a)
end program t
