! vybe-test: fortran/allocation_mold/allocation_mold_copies_scalar_payload
! origin: languages/fortran/tests/fortran/test_allocation_mold.rs
program t
integer, allocatable :: a, b
allocate(b)
b = 17
allocate(a, mold=b)
if ((a) /= 17) then
    print *, "FAIL: want [17] got [", a, "]"
    stop 1
end if
end program t
