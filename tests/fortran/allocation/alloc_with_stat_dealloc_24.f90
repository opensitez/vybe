! vybe-test: fortran/allocation/alloc_with_stat_dealloc_24
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a(:)
integer :: st
allocate(a(2), stat=st)
deallocate(a, stat=st)
end program p
