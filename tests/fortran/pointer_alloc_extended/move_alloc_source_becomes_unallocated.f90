! vybe-test: fortran/pointer_alloc_extended/move_alloc_source_becomes_unallocated
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: src(:), dst(:)
allocate(src(2))
src = [5, 6]
call move_alloc(src, dst)
if ((allocated(src)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(src), "]"
    stop 1
end if
if ((allocated(dst)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", allocated(dst), "]"
    stop 1
end if
end program t
