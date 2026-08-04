! vybe-test: fortran/pointer_alloc_extended/alloc_assign_literal_three_ints
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: v(:)
v = [4, 5, 6]
if ((v(2)) /= 5) then
    print *, "FAIL: want [5] got [", v(2), "]"
    stop 1
end if
if ((size(v)) /= 3) then
    print *, "FAIL: want [3] got [", size(v), "]"
    stop 1
end if
end program t
