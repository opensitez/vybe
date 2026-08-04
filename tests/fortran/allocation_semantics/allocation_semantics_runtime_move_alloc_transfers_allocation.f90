! vybe-test: fortran/allocation_semantics/allocation_semantics_runtime_move_alloc_transfers_allocation
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program t
integer, allocatable :: src(:), dst(:)
allocate(src(2))
src = [8, 9]
call move_alloc(src, dst)
if ((dst(1)) /= 8) then
    print *, "FAIL: want [8] got [", dst(1), "]"
    stop 1
end if
if ((dst(2)) /= 9) then
    print *, "FAIL: want [9] got [", dst(2), "]"
    stop 1
end if
if ((allocated(src)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(src), "]"
    stop 1
end if
if ((allocated(dst)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", allocated(dst), "]"
    stop 1
end if
end program t
