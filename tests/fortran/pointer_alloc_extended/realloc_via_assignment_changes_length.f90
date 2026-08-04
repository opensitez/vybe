! vybe-test: fortran/pointer_alloc_extended/realloc_via_assignment_changes_length
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: seq(:)
seq = [1]
if ((size(seq)) /= 1) then
    print *, "FAIL: want [1] got [", size(seq), "]"
    stop 1
end if
seq = [1, 2, 3, 4, 5]
if ((size(seq)) /= 5) then
    print *, "FAIL: want [5] got [", size(seq), "]"
    stop 1
end if
if ((seq(5)) /= 5) then
    print *, "FAIL: want [5] got [", seq(5), "]"
    stop 1
end if
end program t
