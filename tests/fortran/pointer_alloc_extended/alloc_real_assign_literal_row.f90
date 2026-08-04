! vybe-test: fortran/pointer_alloc_extended/alloc_real_assign_literal_row
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
real, allocatable :: row(:)
row = [0.5, 1.5, 2.5]
if (abs((row(2)) - 1.5) > 1.0e-6) then
    print *, "FAIL: want [1.5] got [", row(2), "]"
    stop 1
end if
if ((size(row)) /= 3) then
    print *, "FAIL: want [3] got [", size(row), "]"
    stop 1
end if
end program t
