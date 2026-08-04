! vybe-test: fortran/allocation_source/allocation_source_copies_scalar_payload
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program t
integer, allocatable :: x
allocate(x, source=5)
if ((x) /= 5) then
    print *, "FAIL: want [5] got [", x, "]"
    stop 1
end if
end program t
