! vybe-test: fortran/pointer_alloc_extended/allocate_stat_reports_success
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: v(:)
integer :: ierr
allocate(v(3), stat=ierr)
if ((ierr) /= 0) then
    print *, "FAIL: want [0] got [", ierr, "]"
    stop 1
end if
if ((size(v)) /= 3) then
    print *, "FAIL: want [3] got [", size(v), "]"
    stop 1
end if
deallocate(v)
end program t
