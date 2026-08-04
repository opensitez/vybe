! vybe-test: fortran/allocation_mold/allocation_mold_copies_array_payload_and_shape
! origin: languages/fortran/tests/fortran/test_allocation_mold.rs
program t
integer, allocatable :: a(:), b(:)
allocate(b(3))
b = [2, 4, 6]
allocate(a, mold=b)
if ((size(a)) /= 3) then
    print *, "FAIL: want [3] got [", size(a), "]"
    stop 1
end if
if ((a(1)) /= 2) then
    print *, "FAIL: want [2] got [", a(1), "]"
    stop 1
end if
if ((a(2)) /= 4) then
    print *, "FAIL: want [4] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 6) then
    print *, "FAIL: want [6] got [", a(3), "]"
    stop 1
end if
end program t
