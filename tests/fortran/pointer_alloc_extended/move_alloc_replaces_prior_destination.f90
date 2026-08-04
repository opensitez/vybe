! vybe-test: fortran/pointer_alloc_extended/move_alloc_replaces_prior_destination
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: fresh(:), old(:)
allocate(fresh(2))
fresh = [100, 200]
old = [9]
call move_alloc(fresh, old)
if ((old(1)) /= 100) then
    print *, "FAIL: want [100] got [", old(1), "]"
    stop 1
end if
if ((size(old)) /= 2) then
    print *, "FAIL: want [2] got [", size(old), "]"
    stop 1
end if
end program t
