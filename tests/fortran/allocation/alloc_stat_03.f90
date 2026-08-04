! vybe-test: fortran/allocation/alloc_stat_03
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3), stat=st)
end program p
