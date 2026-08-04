! vybe-test: fortran/allocation_source/allocation_source_copies_array_payload_to_destination
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program t
integer, allocatable :: a(:)
allocate(a(3), source=[1,2,3])
if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
if ((a(2)) /= 2) then
    print *, "FAIL: want [2] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 3) then
    print *, "FAIL: want [3] got [", a(3), "]"
    stop 1
end if
end program t
